from aiogram import Bot, Dispatcher
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.utils import executor
from keyboards.menu import main_menu

import handlers.menu as menu
import handlers.namebot as proc1
import handlers.qrbot as proc2
import handlers.fallback as proc3

API_TOKEN = ""
    
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

# Регистрируем хендлеры
menu.register_handlers(dp)
proc1.register_handlers(dp)
proc2.register_handlers(dp)
proc3.register_handlers(dp)
# Запуск поллинга
if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
