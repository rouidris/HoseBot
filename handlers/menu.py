from aiogram import types
from aiogram.dispatcher import Dispatcher
from aiogram.dispatcher.filters import Command
from keyboards.menu import main_menu

async def show_menu(message: types.Message):
    await message.answer("Выберите процедуру:", reply_markup=main_menu)


def register_handlers(dp: Dispatcher):
    dp.register_message_handler(show_menu, commands=['start'])