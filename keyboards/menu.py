from aiogram.types import ReplyKeyboardMarkup, KeyboardButton

main_menu = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton("🔧 Процедура 1"), KeyboardButton("📊 Процедура 2")],
        [KeyboardButton("❌ Отмена")]
    ],
    resize_keyboard=True
)
