import os
import pandas as pd
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from barcode.ean import EuropeanArticleNumber13
from barcode.writer import ImageWriter
from PyPDF2 import PdfMerger, PdfReader

from aiogram import types
from aiogram.types import InputFile
from aiogram.dispatcher import Dispatcher, FSMContext
from aiogram.dispatcher.filters import Text
from aiogram.dispatcher.filters.state import State, StatesGroup
from keyboards.menu import main_menu

# Регистрируем шрифт
arial_font_path = 'arial.ttf'
pdfmetrics.registerFont(TTFont('Arial', arial_font_path))

class QRBotStates(StatesGroup):
    waiting_excel = State()
    waiting_pdf = State()

async def start_qrbot(message: types.Message):
    await message.answer('Пожалуйста, отправьте Excel-файл (.xls или .xlsx) для обработки:')
    await QRBotStates.waiting_excel.set()

async def process_excel(message: types.Message, state: FSMContext):
    if not message.document or not message.document.file_name.lower().endswith(('.xls', '.xlsx')):
        await message.answer('Это не Excel-файл. Отправьте файл с расширением .xls или .xlsx.')
        return
    file = await message.document.get_file()
    excel_path = f'tmp_{message.from_user.id}.xlsx'
    await file.download(destination_file=excel_path)
    await state.update_data(excel_path=excel_path)
    await message.answer('Excel-файл получен. Теперь отправьте PDF-файл (.pdf):')
    await QRBotStates.waiting_pdf.set()

async def process_pdf(message: types.Message, state: FSMContext):
    data = await state.get_data()
    excel_path = data.get('excel_path')
    if not message.document or not message.document.file_name.lower().endswith('.pdf'):
        await message.answer('Это не PDF-файл. Отправьте файл с расширением .pdf.')
        return
    file = await message.document.get_file()
    pdf_path = f'tmp_{message.from_user.id}.pdf'
    await file.download(destination_file=pdf_path)

    output_path = f'result_{message.from_user.id}.pdf'
    df = pd.read_excel(excel_path)
    external_pdf = PdfReader(pdf_path)
    external_pages = len(external_pdf.pages)

    merger = PdfMerger()

    # Удаляем старые временные штрихкоды
    for f in os.listdir():
        if f.startswith('barcode_') and f.endswith('.pdf'):
            os.remove(f)
    if os.path.exists('barcode.png'):
        os.remove('barcode.png')

    ext_index = 0
    # Проходим по строкам начиная с третьей (df.iloc[2:])
    for _, row in df.iloc[2:].iterrows():
        name = str(row.get('Unnamed: 2', '')).strip()
        article = str(row.get('Unnamed: 4', '')).strip()
        barcode_num = str(row.get('Unnamed: 6', '')).split(',')[0].strip()

        if not barcode_num.isdigit() or barcode_num == '0':
            continue

        # Генерация штрихкода в PNG
        ean = EuropeanArticleNumber13(barcode_num, writer=ImageWriter())
        ean.save('barcode')  # создаст barcode.png

        # Создание PDF-страницы для штрихкода
        fname = f'barcode_{ext_index}.pdf'
        c = canvas.Canvas(fname, pagesize=(58*mm, 40*mm))
        c.setFont('Arial', 7)
        x, y = 2*mm, 35*mm
        for line in name.split('\n'):
            c.drawString(x, y, line)
            y -= 3.5*mm
        c.drawString(x, y, 'Miss Beauty - будем рады вашему отзыву)')
        y -= 25*mm
        c.drawImage('barcode.png', 0, y, width=60*mm, height=22*mm)
        y -= 25*mm
        c.drawString(x, y, 'Пожалуйста, вскрывайте упаковку')
        y -= 3*mm
        c.drawString(x, y, 'бережно, чтобы подарок вас порадовал!')
        c.save()

        # Удаляем PNG, чтобы не накапливать
        if os.path.exists('barcode.png'):
            os.remove('barcode.png')

        # Добавляем две копии штрихкода
        merger.append(fname)
        merger.append(fname)

        # Добавляем страницу из внешнего PDF
        if ext_index < external_pages:
            # pages=(start, end) добавляет страницы start..end-1, поэтому тут добавляется ровно одна
            merger.append(pdf_path, pages=(ext_index, ext_index+1))
            ext_index += 1

    # Если не было ни одной итерации — сообщаем пользователю
    if ext_index == 0:
        await message.answer('В Excel нет строк с корректными штрихкодами. Попробуйте снова.')
        await state.finish()
        return

    # Сохраняем результирующий PDF
    with open(output_path, 'wb') as f_out:
        merger.write(f_out)
    merger.close()

    # Отправляем результат (читаем как InputFile, чтобы можно было удалить файл сразу)
    doc = InputFile(output_path, filename='result.pdf')
    await message.answer_document(doc)

    # Убираем временные файлы
    for tmp in [excel_path, pdf_path, output_path]:
        if os.path.exists(tmp):
            os.remove(tmp)
    for f in os.listdir():
        if f.startswith('barcode_') and f.endswith('.pdf'):
            os.remove(f)

    await message.answer('Готово! Возвращаемся в меню.', reply_markup=main_menu)
    await state.finish()

def register_handlers(dp: Dispatcher):
    dp.register_message_handler(start_qrbot, Text(equals='📊 Процедура 2'))
    dp.register_message_handler(process_excel, content_types=['document'], state=QRBotStates.waiting_excel)
    dp.register_message_handler(process_pdf, content_types=['document'], state=QRBotStates.waiting_pdf)
