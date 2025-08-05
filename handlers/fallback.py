from aiogram import types
from aiogram.dispatcher import Dispatcher, FSMContext
from aiogram.dispatcher.filters import Text
from keyboards.menu import main_menu

# Обработчик отмены для любого состояния
async def cancel_handler(message: types.Message, state: FSMContext):
    current = await state.get_state()
    if current is None:
        # Если нет активного состояния — просто показываем меню
        await message.answer("Возвращаемся в меню.", reply_markup=main_menu)
        return
    # Сбрасываем состояние
    await state.finish()
    await message.answer("Процедура отменена. Выберите новую команду:", reply_markup=main_menu)

def register_handlers(dp: Dispatcher):
    # срабатывает на кнопку ❌ Отмена в любом состоянии
    dp.register_message_handler(cancel_handler, Text(equals="❌ Отмена"), state="*")
