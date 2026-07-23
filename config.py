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

# Путь к json-ключу сервисного аккаунта Google (для локального запуска —
# файл лежит рядом с bot.py и НЕ попадает в git из-за .gitignore)
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "service_account.json")

# Альтернатива для хостинга (Amvera и т.п.): содержимое json-ключа целиком,
# вставленное как значение одной переменной окружения/секрета. Так как файл
# service_account.json не попадает в git-репозиторий (он в .gitignore), на
# хостинге его физически негде взять из репозитория — вместо этого его
# содержимое вставляется в панели хостинга как секрет с этим именем.
# Если эта переменная задана - используется она, а GOOGLE_CREDENTIALS_FILE
# игнорируется.
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")

# Названия листов внутри таблицы (вкладки снизу в Google Sheets)
GOOGLE_SHEET_WORKSHEET_REG = os.getenv("GOOGLE_SHEET_WORKSHEET_REG", "Регистрация")
GOOGLE_SHEET_WORKSHEET_RECEIPTS = os.getenv("GOOGLE_SHEET_WORKSHEET_RECEIPTS", "Чеки")
GOOGLE_SHEET_WORKSHEET_MODERATORS = os.getenv("GOOGLE_SHEET_WORKSHEET_MODERATORS", "Модераторы")
GOOGLE_SHEET_WORKSHEET_TEXTS = os.getenv("GOOGLE_SHEET_WORKSHEET_TEXTS", "Тексты")

# Папка, куда будут скачиваться фото чеков
RECEIPTS_DIR = os.getenv("RECEIPTS_DIR", "receipts")

# Telegram ID ВЛАДЕЛЬЦЕВ бота, через запятую (например "111111111,222222222").
# Владельцы — это те, кого нельзя удалить из панели: у них всегда есть
# полный доступ, включая управление модераторами. Обычных модераторов
# дальше можно добавлять/удалять прямо из бота (см. пункт меню
# "Модераторы") — они хранятся в Google Таблице, а не здесь.
# Свой ID можно узнать, написав любому боту-помощнику, например @userinfobot.
ADMIN_IDS = {
    int(item.strip())
    for item in os.getenv("ADMIN_IDS", "").split(",")
    if item.strip()
}

if not BOT_TOKEN:
    raise RuntimeError(
        "Не найден BOT_TOKEN. Скопируйте .env.example в .env и вставьте туда "
        "токен, полученный от @BotFather."
    )
