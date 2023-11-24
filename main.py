import logging
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from Token_API import TOKEN_API
from search_name_table import folder_path, get_file_names
from google_func import main_func
from Function_bot import delete_xlsx_files

import logging

# Token Declaration
API_TOKEN = TOKEN_API

# Setting up logging
logging.basicConfig(level=logging.INFO)
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)

# Initializing the bot and dispatcher
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)


# Handler of the start command
@dp.message_handler(commands=['start'])
async def start_command(message: types.Message):
    await message.reply(
        "Вот что я могу:\n"
        f"Отправь мне документ и я его сохраню\n"
        f"/push_table - Отправляем данные в таблицу (функциональна после отправленных документов)\n"
        f"/delete_xlsx - удаляет все xlsx файлы в корневой папке\n"
        f"/check_API - выводит сайт для проверки вкл/выкл API Google Sheets"
    )

# Download the files to the root folder
@dp.message_handler(content_types=types.ContentTypes.DOCUMENT)
async def handle_document(message: types.Document):

    # Getting information about the file
    file_id = message.document.file_id
    file_name = message.document.file_name

    # Saving the file
    file_path = await bot.get_file(file_id)
    file = await bot.download_file(file_path.file_path)

    # Save the file on the computer
    with open(folder_path + file_name, 'wb') as f:
        f.write(file.read())

    # Send the user a message about saving the file
    await message.reply(f"Файл '{file_name}' сохранен.")


# Push files to the table
@dp.message_handler(commands=['push_table'])
async def start_command(message: types.Message):
    # We get the names of xlsx files and pass them to the function
    list_names = get_file_names(folder_path)
    main_func(list_names)
    await message.reply("Данные занесены в таблицу, смотри <link_sheet> ")


# Removing xlsx from the Root folder
@dp.message_handler(commands=['delete_xlsx'])
async def start_command(message: types.Message):
    delete_xlsx_files()
    await message.reply("Все файлы .xlsx удалены из корневой папки")


@dp.message_handler(commands=['check_API'])
async def start_command(message: types.Message):
    await message.reply("Проверь включен ли API <link_API_settings>")

# Launching the bot
if __name__ == '__main__':
    from aiogram import executor

    executor.start_polling(dp, skip_updates=True)