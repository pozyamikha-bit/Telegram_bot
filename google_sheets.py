"""
Вспомогательные функции для работы с Google Таблицей: запись регистраций
и чеков, а также чтение/обновление данных для админ-панели бота.
"""

import datetime
import json
import logging
import time

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

# Заголовки создаются автоматически при первом создании листа. Если лист
# уже существовал ДО перехода на эту версию кода (без заголовка) —
# добавьте эту строку первой строкой в лист вручную в Google Sheets,
# иначе первая строка с данными будет ошибочно принята за заголовок.
REGISTRATION_HEADER = [
    "Дата регистрации", "ФИО", "Магазин", "Телефон", "Telegram ID", "Username",
    "ИНН магазина",
]
RECEIPTS_HEADER = [
    "Дата", "Telegram ID", "Username", "File ID", "Файл", "Статус", "Купон",
    "Комментарий", "Удалён", "№ УПД", "Дата чека",
]
MODERATORS_HEADER = ["Telegram ID", "Username", "Добавил (ID)", "Дата добавления"]

STATUS_PENDING = "на модерации"
STATUS_ACCEPTED = "принят"
STATUS_REJECTED = "отклонён"

TEXTS_HEADER = ["Ключ", "Текст"]

# Импортируемые владельцем списки (админ-панель -> "📥 Импорт"). Каждая
# новая загрузка полностью заменяет предыдущее содержимое листа.
PROMO_ITEMS_HEADER = ["Артикул", "Наименование"]
SALES_REPORT_HEADER = [
    "ИНН", "Наименование", "Артикул", "Количество", "Цена", "№ УПД", "Дата продажи",
    "Дата загрузки",
]

# Тексты-автоответы по умолчанию. Их можно менять прямо из бота
# (админ-панель -> "✏️ Тексты"), тогда значение сохраняется в лист
# "Тексты" и берётся оттуда; если строки для ключа ещё нет — используется
# значение по умолчанию отсюда.
DEFAULT_TEXTS = {
    "rules": (
        "Правила акции:\n"
        "1. Чек должен быть не старше 14 дней.\n"
        "2. На фото должны быть видны дата и название товара.\n"
        "3. Один чек — один купон Ozon."
    ),
    "ask_photo": (
        "Пришлите ОДНО фото чека или УПД. "
        "На фото должно быть видно дату и название товара.\n\n"
        "После фото я попрошу вручную ввести номер УПД и дату чека."
    ),
    "registration_success": "Регистрация успешна! Теперь ты можешь отправлять чеки.",
    "ask_upd_number": "Введите номер УПД, указанный на чеке.",
    "ask_receipt_date": "Введите дату чека в формате ДД.ММ.ГГГГ (например, 21.07.2026).",
    "receipt_received": "Спасибо! Чек принят на модерацию. Ожидайте купон Ozon.",
    "reject_message": "Плохое качество фото чека, просьба повторно зарегистрировать чек",
    "accept_message": "Ваш чек принят! Ваш купон Ozon: {coupon}",
}

# Подписи для меню редактирования текстов в админ-панели.
TEXT_LABELS = {
    "rules": "Текст кнопки «Правила»",
    "ask_photo": "Просьба прислать фото чека",
    "registration_success": "Сообщение после успешной регистрации",
    "ask_upd_number": "Просьба ввести номер УПД (после фото чека)",
    "ask_receipt_date": "Просьба ввести дату чека (после номера УПД)",
    "receipt_received": "Ответ сразу после получения чека",
    "reject_message": "Сообщение при отклонении кнопкой «Быстрое отклонение»",
    "accept_message": "Сообщение при принятии чека (внутри можно оставить {coupon} — вместо него подставится номер купона)",
}

# Московское время: фиксированное смещение UTC+3 (в России нет перехода
# на летнее/зимнее время с 2014 года), поэтому это всегда корректно и не
# зависит от того, в каком часовом поясе запущен сам сервер (хостинг вроде
# Amvera обычно работает в UTC) и не требует установки доп. библиотек.
MOSCOW_TZ = datetime.timezone(datetime.timedelta(hours=3))

_client = None
_spreadsheet = None
_worksheet_cache = {}  # sheet_name -> gspread.Worksheet

# Короткий кэш результатов get_all_records() на лист, чтобы не делать
# отдельный запрос к Google Sheets API на каждое обращение к данным.
# Это и есть главная причина ошибки 429 (RESOURCE_EXHAUSTED /
# RATE_LIMIT_EXCEEDED, лимит 60 запросов на чтение в минуту): без кэша
# один только показ списка чеков в "Истории"/"Модерации" делал отдельный
# запрос на КАЖДЫЙ чек (чтобы найти его регистрацию), и лимит выбирался
# за несколько нажатий кнопок.
_RECORDS_CACHE_TTL_SECONDS = 8
_records_cache = {}  # sheet_name -> (timestamp, records)


def _get_client():
    global _client
    if _client is None:
        if config.GOOGLE_CREDENTIALS_JSON:
            # Хостинг: ключ передан как содержимое переменной окружения
            info = json.loads(config.GOOGLE_CREDENTIALS_JSON)
            creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        else:
            # Локальный запуск: ключ лежит в файле рядом с bot.py
            creds = Credentials.from_service_account_file(
                config.GOOGLE_CREDENTIALS_FILE, scopes=SCOPES
            )
        _client = gspread.authorize(creds)
    return _client


def _get_spreadsheet():
    global _spreadsheet
    if _spreadsheet is None:
        client = _get_client()
        _spreadsheet = client.open_by_key(config.GOOGLE_SHEET_ID)
    return _spreadsheet


def _ensure_header(ws, expected_header):
    """Если в листе уже есть заголовок, но в нём не хватает только новых
    столбцов, добавленных в конце (например, при обновлении кода бота) —
    дописывает их в первую строку автоматически, чтобы не приходилось
    лезть в Google Таблицу руками. Если заголовок отличается как-то иначе
    (другой порядок, опечатки, лишние пустые ячейки в середине) — ничего
    не трогает, чтобы не испортить существующие данные; в этом случае
    нужно поправить вручную (см. check_headers.py)."""
    try:
        actual = ws.row_values(1)
    except Exception:
        logger.exception("Не удалось прочитать заголовок листа '%s'", ws.title)
        return

    if not actual or actual == expected_header:
        return

    n = len(actual)
    if n < len(expected_header) and actual == list(expected_header[:n]):
        missing = expected_header[n:]
        try:
            for i, col_name in enumerate(missing, start=n + 1):
                ws.update_cell(1, i, col_name)
            logger.info(
                "В листе '%s' автоматически дописаны недостающие заголовки: %s",
                ws.title, missing,
            )
        except Exception:
            logger.exception("Не удалось дописать заголовки в лист '%s'", ws.title)


def _get_worksheet(sheet_name: str, header=None):
    if sheet_name in _worksheet_cache:
        return _worksheet_cache[sheet_name]

    spreadsheet = _get_spreadsheet()
    try:
        ws = spreadsheet.worksheet(sheet_name)
        if header:
            _ensure_header(ws, header)
    except gspread.WorksheetNotFound:
        # Если листа с таким названием ещё нет в таблице - создаём его сам
        logger.info("Лист '%s' не найден, создаю новый", sheet_name)
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=10)
        if header:
            ws.append_row(header)
    _worksheet_cache[sheet_name] = ws
    return ws


def _get_records(sheet_name: str, header=None):
    """get_all_records() с коротким кэшем (несколько секунд) — резко
    уменьшает число обращений к Google Sheets API при показе списков."""
    cached = _records_cache.get(sheet_name)
    now = time.time()
    if cached is not None and (now - cached[0]) < _RECORDS_CACHE_TTL_SECONDS:
        return cached[1]

    ws = _get_worksheet(sheet_name, header=header)
    records = ws.get_all_records()
    _records_cache[sheet_name] = (now, records)
    return records


def _invalidate_records_cache(sheet_name: str):
    """Вызывается после любой записи в лист, чтобы следующее чтение сразу
    видело новые данные, а не отдавало устаревшие из кэша."""
    _records_cache.pop(sheet_name, None)


def moscow_now() -> datetime.datetime:
    """Текущие дата и время по Москве."""
    return datetime.datetime.now(MOSCOW_TZ)


def moscow_today() -> datetime.date:
    """Текущая дата по Москве (используется, например, для календаря
    "История" и имени файла отчёта, чтобы не зависеть от часового пояса
    сервера)."""
    return moscow_now().date()


def _moscow_now_str() -> str:
    return moscow_now().strftime("%Y-%m-%d %H:%M:%S")


# ---------- Запись данных (используется в основном сценарии бота) ----------

def append_registration(
    full_name: str, shop: str, phone: str, telegram_id: int, username: str, inn: str = "",
):
    """Добавляет строку с данными регистрации в лист 'Регистрация'."""
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_REG, header=REGISTRATION_HEADER)
    ws.append_row([
        _moscow_now_str(),
        full_name,
        shop,
        phone,
        str(telegram_id),
        username or "",
        inn or "",
    ])
    _invalidate_records_cache(config.GOOGLE_SHEET_WORKSHEET_REG)


def append_receipt(
    telegram_id: int,
    username: str,
    file_id: str,
    file_name: str,
    upd_number: str = "",
    receipt_date: str = "",
):
    """Добавляет строку с данными о присланном чеке в лист 'Чеки'.
    upd_number и receipt_date — номер УПД и дата чека, которые пользователь
    вводит вручную (текстом) сразу после отправки фото."""
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_RECEIPTS, header=RECEIPTS_HEADER)
    ws.append_row([
        _moscow_now_str(),
        str(telegram_id),
        username or "",
        file_id,
        file_name,
        STATUS_PENDING,
        "",
        "",
        "",
        upd_number or "",
        receipt_date or "",
    ])
    _invalidate_records_cache(config.GOOGLE_SHEET_WORKSHEET_RECEIPTS)


# ---------- Чтение и обновление данных (используется в админ-панели) ----------

def get_all_registrations():
    """Возвращает все регистрации списком словарей (для выгрузки отчёта)."""
    records = _get_records(config.GOOGLE_SHEET_WORKSHEET_REG, header=REGISTRATION_HEADER)
    result = []
    for rec in records:
        result.append({
            "date": str(rec.get("Дата регистрации", "")),
            "full_name": str(rec.get("ФИО", "")),
            "shop": str(rec.get("Магазин", "")),
            "phone": str(rec.get("Телефон", "")),
            "telegram_id": str(rec.get("Telegram ID", "")),
            "username": str(rec.get("Username", "")),
            "inn": str(rec.get("ИНН магазина", "")),
        })
    return result


def get_registration_by_telegram_id(telegram_id):
    """Возвращает последнюю регистрацию пользователя с данным Telegram ID
    в виде {"full_name", "shop", "phone", "inn"} или None, если не найдено."""
    records = _get_records(config.GOOGLE_SHEET_WORKSHEET_REG, header=REGISTRATION_HEADER)
    telegram_id = str(telegram_id)
    for rec in reversed(records):
        if str(rec.get("Telegram ID", "")) == telegram_id:
            return {
                "full_name": str(rec.get("ФИО", "")),
                "shop": str(rec.get("Магазин", "")),
                "phone": str(rec.get("Телефон", "")),
                "inn": str(rec.get("ИНН магазина", "")),
            }
    return None


def get_receipts():
    """Возвращает все чеки списком словарей с ключом 'row' — номером строки
    в таблице (нужен, чтобы потом обновить именно эту строку)."""
    records = _get_records(config.GOOGLE_SHEET_WORKSHEET_RECEIPTS, header=RECEIPTS_HEADER)
    result = []
    for i, rec in enumerate(records):
        result.append({
            "row": i + 2,  # +2: строка 1 - заголовок, gspread индексирует с 1
            "date": str(rec.get("Дата", "")),
            "telegram_id": str(rec.get("Telegram ID", "")),
            "username": str(rec.get("Username", "")),
            "file_id": str(rec.get("File ID", "")),
            "file_name": str(rec.get("Файл", "")),
            "status": str(rec.get("Статус", "")),
            "coupon": str(rec.get("Купон", "")),
            "comment": str(rec.get("Комментарий", "")),
            "deleted": str(rec.get("Удалён", "")).strip().lower() in ("да", "yes", "true", "1"),
            "upd_number": str(rec.get("№ УПД", "")),
            "receipt_date": str(rec.get("Дата чека", "")),
        })
    return result


def get_receipts_by_date(date_str: str, include_deleted: bool = False):
    """date_str в формате YYYY-MM-DD. Возвращает чеки за эту дату, любой
    статус. Помеченные удалёнными по умолчанию скрыты."""
    return [
        r for r in get_receipts()
        if r["date"].startswith(date_str) and (include_deleted or not r["deleted"])
    ]


def get_pending_receipts():
    """Возвращает чеки со статусом 'на модерации' (без помеченных удалёнными)."""
    return [r for r in get_receipts() if r["status"] == STATUS_PENDING and not r["deleted"]]


def get_receipts_by_telegram_id(telegram_id, include_deleted: bool = True):
    """Возвращает все чеки конкретного пользователя (для его личного отчёта
    "Загруженные чеки") — по умолчанию включая помеченные удалёнными, чтобы
    пользователь видел полную историю своих загрузок."""
    telegram_id = str(telegram_id)
    return [
        r for r in get_receipts()
        if r["telegram_id"] == telegram_id and (include_deleted or not r["deleted"])
    ]


def get_receipt_by_row(row_number: int):
    for r in get_receipts():
        if r["row"] == row_number:
            return r
    return None


def mark_receipt_deleted(row_number: int, deleted: bool = True):
    """Помечает чек как удалённый (или снимает пометку) — сама строка и
    все данные в ней остаются в таблице, ничего не стирается."""
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_RECEIPTS, header=RECEIPTS_HEADER)
    deleted_col = RECEIPTS_HEADER.index("Удалён") + 1
    ws.update_cell(row_number, deleted_col, "да" if deleted else "")
    _invalidate_records_cache(config.GOOGLE_SHEET_WORKSHEET_RECEIPTS)



def update_receipt_status(row_number: int, status: str, coupon: str = None, comment: str = None):
    """Обновляет статус чека и, опционально, номер купона и/или комментарий
    модератора (например, причину отклонения) по номеру строки."""
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_RECEIPTS, header=RECEIPTS_HEADER)
    status_col = RECEIPTS_HEADER.index("Статус") + 1
    ws.update_cell(row_number, status_col, status)
    if coupon is not None:
        coupon_col = RECEIPTS_HEADER.index("Купон") + 1
        ws.update_cell(row_number, coupon_col, coupon)
    if comment is not None:
        comment_col = RECEIPTS_HEADER.index("Комментарий") + 1
        ws.update_cell(row_number, comment_col, comment)
    _invalidate_records_cache(config.GOOGLE_SHEET_WORKSHEET_RECEIPTS)


# ---------- Управление модераторами (доступно только владельцам, ADMIN_IDS) ----------

def get_moderators():
    """Возвращает список модераторов вида {row, telegram_id, username, added_by, date}."""
    records = _get_records(config.GOOGLE_SHEET_WORKSHEET_MODERATORS, header=MODERATORS_HEADER)
    result = []
    for i, rec in enumerate(records):
        result.append({
            "row": i + 2,
            "telegram_id": str(rec.get("Telegram ID", "")),
            "username": str(rec.get("Username", "")),
            "added_by": str(rec.get("Добавил (ID)", "")),
            "date": str(rec.get("Дата добавления", "")),
        })
    return result


def get_moderator_ids():
    """Множество Telegram ID всех модераторов (без учёта владельцев из ADMIN_IDS)."""
    ids = set()
    for m in get_moderators():
        try:
            ids.add(int(m["telegram_id"]))
        except ValueError:
            continue
    return ids


def add_moderator(telegram_id: int, username: str, added_by: int):
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_MODERATORS, header=MODERATORS_HEADER)
    ws.append_row([
        str(telegram_id),
        username or "",
        str(added_by),
        _moscow_now_str(),
    ])
    _invalidate_records_cache(config.GOOGLE_SHEET_WORKSHEET_MODERATORS)


def remove_moderator(row_number: int):
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_MODERATORS, header=MODERATORS_HEADER)
    ws.delete_rows(row_number)
    _invalidate_records_cache(config.GOOGLE_SHEET_WORKSHEET_MODERATORS)


# ---------- Редактируемые тексты-автоответы (доступно владельцам) ----------

def get_all_texts():
    """Возвращает словарь key -> текущий текст: из таблицы, если он там
    задан и не пуст, иначе значение по умолчанию из DEFAULT_TEXTS."""
    records = _get_records(config.GOOGLE_SHEET_WORKSHEET_TEXTS, header=TEXTS_HEADER)
    overrides = {str(rec.get("Ключ", "")): str(rec.get("Текст", "")) for rec in records}

    result = dict(DEFAULT_TEXTS)
    for key, value in overrides.items():
        if key in result and value:
            result[key] = value
    return result


def get_text(key: str) -> str:
    return get_all_texts().get(key, DEFAULT_TEXTS.get(key, ""))


def set_text(key: str, value: str):
    """Сохраняет новый текст для ключа: обновляет существующую строку
    в листе "Тексты", либо добавляет новую, если ключа там ещё не было."""
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_TEXTS, header=TEXTS_HEADER)
    records = ws.get_all_records()
    try:
        for i, rec in enumerate(records):
            if str(rec.get("Ключ", "")) == key:
                ws.update_cell(i + 2, 2, value)
                return
        ws.append_row([key, value])
    finally:
        _invalidate_records_cache(config.GOOGLE_SHEET_WORKSHEET_TEXTS)


# ---------- Импорт списков владельцем (акционные позиции / продажи) ----------

def _replace_sheet_data(sheet_name: str, header: list, rows: list):
    """Полностью заменяет содержимое листа: старые данные стираются,
    записывается новый заголовок и переданные строки. Используется для
    "Импорта" в админ-панели, где каждая новая загрузка файла заменяет
    предыдущий список целиком."""
    ws = _get_worksheet(sheet_name, header=header)
    ws.clear()
    ws.update("A1", [header] + [list(row) for row in rows])
    _invalidate_records_cache(sheet_name)


def import_promo_items(rows):
    """rows: список [артикул, наименование]. Полностью заменяет лист
    "Акционные позиции"."""
    _replace_sheet_data(config.GOOGLE_SHEET_WORKSHEET_PROMO_ITEMS, PROMO_ITEMS_HEADER, rows)


def get_promo_items():
    records = _get_records(config.GOOGLE_SHEET_WORKSHEET_PROMO_ITEMS, header=PROMO_ITEMS_HEADER)
    return [
        {
            "article": str(rec.get("Артикул", "")),
            "name": str(rec.get("Наименование", "")),
        }
        for rec in records
    ]


def import_sales_report(rows):
    """rows: список [ИНН, наименование, артикул, количество, цена,
    № УПД, дата продажи]. В отличие от import_promo_items — НЕ заменяет
    предыдущие загрузки, а добавляет новые строки к уже имеющимся
    (история копится), проставляя каждой строке дату и время именно
    этой загрузки (по Москве) — чтобы потом можно было понять, из
    какого именно импорта пришла строка."""
    ws = _get_worksheet(config.GOOGLE_SHEET_WORKSHEET_SALES_REPORT, header=SALES_REPORT_HEADER)
    upload_ts = _moscow_now_str()
    values = [list(row) + [upload_ts] for row in rows]
    if values:
        ws.append_rows(values)
    _invalidate_records_cache(config.GOOGLE_SHEET_WORKSHEET_SALES_REPORT)


def get_sales_report():
    records = _get_records(config.GOOGLE_SHEET_WORKSHEET_SALES_REPORT, header=SALES_REPORT_HEADER)
    return [
        {
            "inn": str(rec.get("ИНН", "")),
            "name": str(rec.get("Наименование", "")),
            "article": str(rec.get("Артикул", "")),
            "quantity": str(rec.get("Количество", "")),
            "price": str(rec.get("Цена", "")),
            "upd_number": str(rec.get("№ УПД", "")),
            "sale_date": str(rec.get("Дата продажи", "")),
            "upload_date": str(rec.get("Дата загрузки", "")),
        }
        for rec in records
    ]
