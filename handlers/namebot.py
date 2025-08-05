import os
import pandas as pd
from aiogram import types
from aiogram.dispatcher import Dispatcher, FSMContext
from aiogram.dispatcher.filters import Text
from aiogram.dispatcher.filters.state import State, StatesGroup
from keyboards.menu import main_menu

# Состояния для загрузки Excel-файла
class Proc1States(StatesGroup):
    waiting_excel = State()

async def start_proc1(message: types.Message):
    await message.answer("Пожалуйста, отправьте Excel-файл (.xls или .xlsx) для анализа:")
    await Proc1States.waiting_excel.set()

async def process_excel(message: types.Message, state: FSMContext):
    if not message.document or not message.document.file_name.lower().endswith(('.xls', '.xlsx')):
        await message.answer("Это не Excel-файл. Отправьте .xls или .xlsx.")
        return
    file = await message.document.get_file()
    xlsx_path = f'tmp1_{message.from_user.id}.xlsx'
    await file.download(destination_file=xlsx_path)

    # Читаем excel
    df = pd.read_excel(xlsx_path, sheet_name="Лист подбора", header=3)
    df.columns = ["Фото", "Бренд", "Наименование", "Артикул продавца", "Цвет", "Стикер", "Баркод"]
    df = df[df["Бренд"] != "Бренд"].dropna(subset=["Наименование"])
    df["Бренд"] = df["Бренд"].fillna("")

    # Группируем
    grouped = df.groupby(["Бренд", "Наименование"]).size().reset_index(name="Количество")
    grouped = grouped.sort_values(by="Количество", ascending=False)

    # Отправляем результаты
    total = grouped["Количество"].sum()
    for _, row in grouped.iterrows():
        brand = row["Бренд"]
        name = row["Наименование"]
        count = row["Количество"]
        await message.answer(f"{brand} — {name}: {count} штук")

    await message.answer(f"Общее количество товаров: {total} шт.", reply_markup=main_menu)

    # Очистка и сброс состояния
    await state.finish()
    if os.path.exists(xlsx_path): os.remove(xlsx_path)


def register_handlers(dp: Dispatcher):
    dp.register_message_handler(start_proc1, Text(equals="🔧 Процедура 1"))
    dp.register_message_handler(process_excel, content_types=['document'], state=Proc1States.waiting_excel)