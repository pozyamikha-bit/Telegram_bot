"""
Вспомогательные функции для записи данных в Google Таблицу.

Используется связка gspread + google-auth и сервисный аккаунт Google
(см. инструкцию в README.md, раздел про настройку Google Sheets API).
"""

import datetime
import logging

import gspread
from google.oauth2.service_account import Credentials

import config

logger = logging.getLogger(__name__)

# Права, которые запрашиваем у Google: доступ к таблицам и к диску
# (доступ к диску нужен, чтобы gspread мог найти таблицу по ID)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_client = None


def _get_client():
    global _client
    if _client is None:
        creds = Credentials.from_service_account_file(
            config.GOOGLE_CREDENTIALS_FILE, scopes=SCOPES
        )
        _client = gspread.authorize(creds)
    return _client


def _get_worksheet(sheet_name: str):
    client = _get_client()
    spreadsheet = client.open_by_key(config.GOOGLE_SHEET_ID)
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        # Если листа с таким названием ещё нет в таблице - создаём его сам
        logger.info("Лист '%s' не найден, создаю новый", sheet_name)
        return spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=10)


def append_registration(full_name: str, shop: str, phone: str, telegram_id: int, username: str):
    """Добавляет строку с данными регистрации в лист 'Регистрация'."""
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_REG)
    ws.append_row([
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        full_name,
        shop,
        phone,
        str(telegram_id),
        username or "",
    ])


def append_receipt(telegram_id: int, username: str, file_id: str, file_name: str):
    """Добавляет строку с данными о присланном чеке в лист 'Чеки'."""
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_RECEIPTS)
    ws.append_row([
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        str(telegram_id),
        username or "",
        file_id,
        file_name,
        "на модерации",
    ])
