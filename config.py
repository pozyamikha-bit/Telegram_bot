"""
Настройки бота.

Все секретные значения (токен, ID таблицы) берутся из файла .env,
который НЕ должен попадать в открытый доступ (например, в публичный
git-репозиторий). Скопируйте .env.example в .env и заполните своими
значениями.
"""

import os

from dotenv import load_dotenv

load_dotenv()

# Токен бота, полученный от @BotFather
BOT_TOKEN = os.getenv("BOT_TOKEN")

# ID Google-таблицы. Это часть ссылки на таблицу между /d/ и /edit,
# например для https://docs.google.com/spreadsheets/d/1AbCDefGhIjKl/edit
# ID будет 1AbCDefGhIjKl
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")

# Путь к json-ключу сервисного аккаунта Google
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "service_account.json")

# Названия листов внутри таблицы (вкладки снизу в Google Sheets)
GOOGLE_SHEET_WORKSHEET_REG = os.getenv("GOOGLE_SHEET_WORKSHEET_REG", "Регистрация")
GOOGLE_SHEET_WORKSHEET_RECEIPTS = os.getenv("GOOGLE_SHEET_WORKSHEET_RECEIPTS", "Чеки")

# Папка, куда будут скачиваться фото чеков
RECEIPTS_DIR = os.getenv("RECEIPTS_DIR", "receipts")

if not BOT_TOKEN:
    raise RuntimeError(
        "Не найден BOT_TOKEN. Скопируйте .env.example в .env и вставьте туда "
        "токен, полученный от @BotFather."
    )
