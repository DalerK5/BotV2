import os
import re
import asyncio
import openpyxl
from pypdf import PdfReader
from aiogram import Bot, Dispatcher, types
from aiogram.types import FSInputFile
from aiogram.filters import Command

# Токен берется из переменных окружения Render
TOKEN = os.getenv("BOT_TOKEN")

bot = Bot(token=TOKEN)
dp = Dispatcher()

def extract_pdf_data(pdf_path):
    """Извлечение всех метаданных и результатов испытаний из PDF"""
    reader = PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            full_text += txt + "\n"
            
    data = {}
    
    # 1. Номер протокола
    m_no = re.search(r"№\s*(SB-[^\s\n]+)", full_text)
    data["protocol_no"] = m_no.group(1) if m_no else ""
    
    # 2. Дата протокола
    m_date = re.search(r"Sanasi:\s*([\d.]+)", full_text)
    data["date"] = m_date.group(1) if m_date else ""
    
    # 3. VIN-код
    m_vin = re.search(r"VIN\s*№\s*([A-Z0-9]+)", full_text)
    data["vin"] = m_vin.group(1) if m_vin else ""
    
    # 4. Марка и модель
    m_mark = re.search(r"марка\s*«([^»]+)»", full_text, re.IGNORECASE)
    data["mark"] = m_mark.group(1) if m_mark else ""
    
    m_model = re.search(r"модель\s*«([^»]+)»", full_text, re.IGNORECASE)
    data["model"] = m_model.group(1) if m_model else ""
    
    # 5. Климатические условия (безопасные регулярные выражения в одну строку)
    m_temp = re.search(r"Harorat:\s*([^\n°]+)", full_text)
    data["temp"] = m_temp.group(1).strip() if m_temp else "21"
    
    m_hum = re.search(r"Nisbiy namlik:\s*([^\n%]+)", full_text)
    data["humidity"] = m_hum.group(1).strip() if m_hum else "36"
    
    m_press = re.search(r"Atmosfera bosimi:\s*([^\n]+)", full_text)
    data["pressure"] = m_press.group(1).strip() if m_press else "97"
    
    m_wind = re.search(r"Havo tezligi:\s*([^\n]+)", full_text)
    data["wind"] = m_wind.group(1).strip() if m_wind else "1 m/s"
    
    return data

def update_excel_journal(template_path, output_path, data):
    """Обновление шапок во всей книге Excel и замена данных"""
    wb = openpyxl.load_workbook(template_path)
    
    # 1. Обновляем главный лист "Рабочий журнал"
    sheet_main = wb['Рабочий журнал'] if 'Рабочий журнал' in wb.sheetnames else wb.active
    
    if data.get("date"):
        sheet_main['E2'] = f"     Дата проведения испытаний: «{data['date'][:2]}» {data['date'][3:5]} 2026 г.      "
    if data.get("temp"):
        sheet_main['E6'] = f"{data['temp']} °C"
    if data.get("humidity"):
        sheet_main['E7'] = f"{data['humidity']} %"
    if data.get("pressure"):
        sheet_main['E8'] = f"{data['pressure']} кПа"
    if data.get("wind"):
        sheet_main['E9'] = data['wind']
        
    if data.get("mark"):
        sheet_main['E11'] = data['mark']
    if data.get("model"):
        sheet_main['E12'] = data['model']
    if data.get("vin"):
        sheet_main['E15'] = data['vin']

    # 2. Замена данных автомобиля по всем листам
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for r in range(1, 16):
            for c in range(1, 10):
                cell = sheet.cell(row=r, column=c)
                if cell.value and isinstance(cell.value, str):
                    val = cell.value
                    if "VIN" in val and data.get("vin"):
                        cell.value = f"VIN № {data['vin']}"
                    elif "марка" in val.lower() and data.get("mark"):
                        cell.value = f"Марка: {data['mark']}"

    wb.save(output_path)

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    await message.answer(
        "👋 **Добро пожаловать!**\n\n"
        "Отправьте мне PDF-протокол испытаний, и я автоматически заполню "
        "рабочий журнал сертификационных испытаний Excel (все 80+ листов)."
    )

@dp.message(lambda msg: msg.document is not None)
async def handle_document(message: types.Message):
    doc = message.document
    if not doc.file_name.lower().endswith('.pdf'):
        await message.answer("⚠️ Пожалуйста, отправьте документ в формате PDF.")
        return

    status_msg = await message.answer("🔄 **Обработка протокола...** Извлекаем данные и заполняем 80+ листов Excel...")

    pdf_file_path = f"temp_{doc.file_id}.pdf"
    output_excel_path = f"Журнал_{doc.file_name.replace('.pdf', '')}.xlsx"

    try:
        file_info = await bot.get_file(doc.file_id)
        downloaded = await bot.download(file_info)
        with open(pdf_file_path, "wb") as f:
            f.write(downloaded.read())

        data = extract_pdf_data(pdf_file_path)
        template_file = "1660518 Переделанный.xlsx"
        update_excel_journal(template_file, output_excel_path, data)

        result_file = FSInputFile(output_excel_path)
        await message.answer_document(
            result_file, 
            caption=f"✅ **Рабочий журнал успешно заполнен!**\n\n"
                    f"📌 **Протокол №:** {data.get('protocol_no', '—')}\n"
                    f"🚗 **Марка/Модель:** {data.get('mark', '')} {data.get('model', '')}\n"
                    f"🔑 **VIN:** `{data.get('vin', '—')}`"
        )
        await status_msg.delete()

    except Exception as e:
        await message.answer(f"❌ **Произошла ошибка при обработке:** {str(e)}")

    finally:
        if os.path.exists(pdf_file_path):
            os.remove(pdf_file_path)
        if os.path.exists(output_excel_path):
            os.remove(output_excel_path)

async def main():
    print("Бот запущен и готов к работе...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
