"""
Бот регистрации участников акции + приём чеков/УПД + админ-панель.

Версия под aiogram 3.x (актуальная ветка библиотеки; именно её ожидает
и платформа Amvera при сборке — см. README.md).

Сценарий для обычного пользователя:
  /start -> ФИО -> Магазин -> Телефон (кнопкой) -> запись в Google Таблицу
  -> главное меню -> "Отправить чек" -> фото -> запись в Google Таблицу

Админ-панель (/admin, доступно владельцам из ADMIN_IDS и модераторам,
добавленным через саму панель):
  - "История" — календарь, по выбранной дате список всех чеков за неё
    с инфо об отправителе и фото.
  - "Модерация" — список чеков со статусом "на модерации"; для каждого
    можно принять (после чего запрашивается номер купона и он
    отправляется пользователю) или отклонить (пользователь уведомляется).
  - "Модераторы" (только для владельцев из ADMIN_IDS) — добавление и
    удаление модераторов.
  - "Отчёт" — Excel-файл со всеми чеками и регистрациями.
"""

import asyncio
import calendar
import io
import logging
import os
import re
from datetime import date

import openpyxl
from openpyxl.utils import get_column_letter

from aiogram import Bot, Dispatcher, F
from aiogram.filters import BaseFilter, Command, CommandStart, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    InlineKeyboardButton,
    Message,
)
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder

import config
import google_sheets

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)

bot = Bot(token=config.BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

os.makedirs(config.RECEIPTS_DIR, exist_ok=True)


class Form(StatesGroup):
    name = State()
    shop = State()
    inn = State()
    phone = State()
    waiting_photo = State()
    waiting_upd_number = State()
    waiting_receipt_date = State()


class AdminForm(StatesGroup):
    waiting_coupon = State()
    waiting_moderator_id = State()
    waiting_text_value = State()
    waiting_reject_reason = State()
    waiting_promo_file = State()
    waiting_sales_file = State()
    waiting_sales_period = State()
    waiting_auto_moderation_period = State()


# ---------- Проверка прав доступа ----------

def _is_owner(user_id: int) -> bool:
    return user_id in config.ADMIN_IDS


async def _is_moderator(user_id: int) -> bool:
    if _is_owner(user_id):
        return True
    try:
        return user_id in google_sheets.get_moderator_ids()
    except Exception:
        logger.exception("Не удалось проверить список модераторов")
        return False


class IsOwner(BaseFilter):
    async def __call__(self, event) -> bool:
        return _is_owner(event.from_user.id)


class IsModerator(BaseFilter):
    async def __call__(self, event) -> bool:
        return await _is_moderator(event.from_user.id)


# ---------- Клавиатуры ----------

BTN_SEND_RECEIPT = "📥 Отправить чек / УПД"
BTN_MY_RECEIPTS = "📄 Загруженные чеки"
BTN_RULES = "❓ Правила"

ADMIN_BTN_HISTORY = "🗓 История"
ADMIN_BTN_MODERATION = "🛠 Модерация"
ADMIN_BTN_MODERATORS = "👥 Модераторы"
ADMIN_BTN_REPORT = "📊 Отчёты"
ADMIN_BTN_TEXTS = "✏️ Тексты"
ADMIN_BTN_IMPORT = "📥 Импорт"
ADMIN_BTN_EXIT = "⬅️ Выйти из панели"

_main_menu_builder = ReplyKeyboardBuilder()
_main_menu_builder.button(text=BTN_SEND_RECEIPT)
_main_menu_builder.button(text=BTN_MY_RECEIPTS)
_main_menu_builder.button(text=BTN_RULES)
_main_menu_builder.adjust(1)
main_menu = _main_menu_builder.as_markup(resize_keyboard=True)

_moderator_menu_builder = ReplyKeyboardBuilder()
_moderator_menu_builder.button(text=ADMIN_BTN_HISTORY)
_moderator_menu_builder.button(text=ADMIN_BTN_MODERATION)
_moderator_menu_builder.button(text=ADMIN_BTN_REPORT)
_moderator_menu_builder.button(text=ADMIN_BTN_EXIT)
_moderator_menu_builder.adjust(2, 2)
moderator_menu = _moderator_menu_builder.as_markup(resize_keyboard=True)

_owner_menu_builder = ReplyKeyboardBuilder()
_owner_menu_builder.button(text=ADMIN_BTN_HISTORY)
_owner_menu_builder.button(text=ADMIN_BTN_MODERATION)
_owner_menu_builder.button(text=ADMIN_BTN_REPORT)
_owner_menu_builder.button(text=ADMIN_BTN_MODERATORS)
_owner_menu_builder.button(text=ADMIN_BTN_TEXTS)
_owner_menu_builder.button(text=ADMIN_BTN_IMPORT)
_owner_menu_builder.button(text=ADMIN_BTN_EXIT)
_owner_menu_builder.adjust(2, 2, 2, 1)
owner_menu = _owner_menu_builder.as_markup(resize_keyboard=True)


def _menu_for(user_id: int):
    return owner_menu if _is_owner(user_id) else moderator_menu


# ---------- Регистрация ----------

@dp.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(Form.name)
    await message.answer(
        "Здравствуйте! Давайте зарегистрируем вас в программе.\n\n"
        "Для начала напишите, пожалуйста, ваше ФИО."
    )


# ---------- Вход в админ-панель (регистрируется рано, чтобы перебивать FSM) ----------

@dp.message(Command("admin"), IsModerator())
async def cmd_admin(message: Message, state: FSMContext):
    await state.clear()
    if _is_owner(message.from_user.id):
        await message.answer("Админ-панель.", reply_markup=owner_menu)
    else:
        await message.answer("Панель модератора.", reply_markup=moderator_menu)


@dp.message(F.text == ADMIN_BTN_EXIT, IsModerator())
async def admin_exit(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Вышли из панели.", reply_markup=main_menu)


# ---------- Регистрация (продолжение сценария) ----------

@dp.message(StateFilter(Form.name))
async def process_name(message: Message, state: FSMContext):
    await state.update_data(name=message.text)
    await state.set_state(Form.shop)
    await message.answer("Отлично! Теперь укажите название вашего магазина.")


@dp.message(StateFilter(Form.shop))
async def process_shop(message: Message, state: FSMContext):
    await state.update_data(shop=message.text)
    await state.set_state(Form.inn)
    await message.answer(
        "Теперь укажите ИНН магазина (10 цифр — для организации, "
        "или 12 цифр — для ИП)."
    )


_INN_RE = re.compile(r"^\d{10}$|^\d{12}$")


@dp.message(StateFilter(Form.inn), F.text)
async def process_inn(message: Message, state: FSMContext):
    inn = message.text.strip()

    if not _INN_RE.match(inn):
        await message.answer(
            "ИНН должен состоять только из цифр — 10 цифр для организации "
            "или 12 цифр для ИП. Введите ИНН ещё раз."
        )
        return

    await state.update_data(inn=inn)

    kb_builder = ReplyKeyboardBuilder()
    kb_builder.button(text="Поделиться контактом", request_contact=True)
    keyboard = kb_builder.as_markup(resize_keyboard=True, one_time_keyboard=True)

    await state.set_state(Form.phone)
    await message.answer(
        "И последнее: поделись номером телефона для связи.",
        reply_markup=keyboard,
    )


@dp.message(StateFilter(Form.inn))
async def wrong_content_for_inn(message: Message):
    await message.answer("Пожалуйста, введите ИНН магазина текстом (только цифры).")


@dp.message(StateFilter(Form.phone), F.contact)
async def process_phone(message: Message, state: FSMContext):
    data = await state.get_data()
    full_name = data.get("name", "")
    shop = data.get("shop", "")
    inn = data.get("inn", "")
    phone = message.contact.phone_number

    try:
        google_sheets.append_registration(
            full_name=full_name,
            shop=shop,
            inn=inn,
            phone=phone,
            telegram_id=message.from_user.id,
            username=message.from_user.username,
        )
    except Exception:
        logger.exception("Не удалось записать регистрацию в Google Таблицу")
        await message.answer(
            "Регистрация прошла, но возникла ошибка при записи в таблицу. "
            "Сообщите об этом администратору."
        )

    await state.clear()
    await message.answer(
        google_sheets.get_text("registration_success"),
        reply_markup=main_menu,
    )


# ---------- Правила ----------

@dp.message(F.text == BTN_RULES)
async def show_rules(message: Message):
    await message.answer(google_sheets.get_text("rules"))


# ---------- Приём чеков ----------

@dp.message(F.text == BTN_SEND_RECEIPT)
async def send_receipt_start(message: Message, state: FSMContext):
    await state.set_state(Form.waiting_photo)
    await message.answer(google_sheets.get_text("ask_photo"))


@dp.message(StateFilter(Form.waiting_photo), F.photo)
async def process_photo(message: Message, state: FSMContext):
    photo = message.photo[-1]  # фото в максимальном качестве
    file_id = photo.file_id

    filename = f"{message.from_user.id}_{message.date.strftime('%Y%m%d_%H%M%S')}.jpg"
    filepath = os.path.join(config.RECEIPTS_DIR, filename)

    try:
        await bot.download(file_id, destination=filepath)
    except Exception:
        logger.exception("Не удалось скачать фото чека")
        filepath = "(не скачалось)"

    # Фото ещё не сохраняем в таблицу — сначала нужно получить от
    # пользователя номер УПД и дату чека вручную, чтобы записать всё
    # одной строкой.
    await state.update_data(photo_file_id=file_id, photo_file_name=filepath)
    await state.set_state(Form.waiting_upd_number)
    await message.answer(google_sheets.get_text("ask_upd_number"))


@dp.message(StateFilter(Form.waiting_photo))
async def wrong_content_for_photo(message: Message):
    await message.answer("Пожалуйста, пришлите именно фото (как изображение, не файлом).")


@dp.message(StateFilter(Form.waiting_upd_number), F.text)
async def process_upd_number(message: Message, state: FSMContext):
    await state.update_data(upd_number=message.text.strip())
    await state.set_state(Form.waiting_receipt_date)
    await message.answer(google_sheets.get_text("ask_receipt_date"))


@dp.message(StateFilter(Form.waiting_upd_number))
async def wrong_content_for_upd_number(message: Message):
    await message.answer("Пожалуйста, введите номер УПД текстом (одним сообщением).")


_DATE_LIKE_RE = re.compile(r"^\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}$")


@dp.message(StateFilter(Form.waiting_receipt_date), F.text)
async def process_receipt_date(message: Message, state: FSMContext):
    receipt_date = message.text.strip()

    if not _DATE_LIKE_RE.match(receipt_date):
        await message.answer(
            "Похоже, это не дата. Введите дату чека в формате ДД.ММ.ГГГГ, "
            "например 21.07.2026."
        )
        return

    data = await state.get_data()

    try:
        google_sheets.append_receipt(
            telegram_id=message.from_user.id,
            username=message.from_user.username,
            file_id=data.get("photo_file_id", ""),
            file_name=data.get("photo_file_name", ""),
            upd_number=data.get("upd_number", ""),
            receipt_date=receipt_date,
        )
    except Exception:
        logger.exception("Не удалось записать чек в Google Таблицу")

    await state.clear()
    await message.answer(
        google_sheets.get_text("receipt_received"),
        reply_markup=main_menu,
    )


@dp.message(StateFilter(Form.waiting_receipt_date))
async def wrong_content_for_receipt_date(message: Message):
    await message.answer("Пожалуйста, введите дату чека текстом, в формате ДД.ММ.ГГГГ.")


# ---------- Загруженные чеки (для самого пользователя) ----------

@dp.message(F.text == BTN_MY_RECEIPTS)
async def my_receipts(message: Message):
    await message.answer("Формирую отчёт по вашим чекам, секунду...")
    try:
        receipts = google_sheets.get_receipts_by_telegram_id(message.from_user.id)
    except Exception:
        logger.exception("Не удалось получить чеки пользователя из Google Таблицы")
        await message.answer("Не получилось сформировать отчёт, попробуйте позже.")
        return

    if not receipts:
        await message.answer("Вы ещё не загружали ни одного чека.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Мои чеки"
    headers = ["Дата загрузки", "№ УПД", "Дата чека", "Статус", "Купон", "Комментарий"]
    ws.append(headers)
    for r in receipts:
        ws.append([
            r["date"],
            r["upd_number"],
            r["receipt_date"],
            r["status"],
            r["coupon"],
            r["comment"],
        ])
    _autosize(ws, headers)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"moi_cheki_{message.from_user.id}.xlsx"
    await message.answer_document(BufferedInputFile(buf.read(), filename=filename))


# ---------- Общая карточка чека (используется и в Истории, и в Модерации) ----------

async def _send_receipt_card(
    message: Message,
    receipt: dict,
    with_moderation_buttons: bool = False,
    with_history_buttons: bool = False,
):
    reg = google_sheets.get_registration_by_telegram_id(receipt["telegram_id"])

    lines = []
    if receipt.get("deleted"):
        lines.append("🗑 Помечен как удалённый")
    if reg:
        lines.append(f"ФИО: {reg['full_name']}")
        lines.append(f"Магазин: {reg['shop']}")
        if reg.get("inn"):
            lines.append(f"ИНН магазина: {reg['inn']}")
        lines.append(f"Телефон: {reg['phone']}")
    else:
        lines.append("Регистрация не найдена (данные могли не сохраниться).")

    lines.append(f"Telegram ID: {receipt['telegram_id']}")
    if receipt["username"]:
        lines.append(f"Username: @{receipt['username']}")
    lines.append(f"Дата загрузки: {receipt['date']}")
    if receipt.get("upd_number"):
        lines.append(f"№ УПД (введён пользователем): {receipt['upd_number']}")
    if receipt.get("receipt_date"):
        lines.append(f"Дата чека (введена пользователем): {receipt['receipt_date']}")
    lines.append(f"Статус: {receipt['status']}")
    if receipt.get("coupon"):
        lines.append(f"Купон: {receipt['coupon']}")
    if receipt.get("comment"):
        lines.append(f"Комментарий: {receipt['comment']}")

    caption = "\n".join(lines)

    reply_markup = None
    if with_moderation_buttons:
        kb = InlineKeyboardBuilder()
        kb.row(InlineKeyboardButton(text="✅ Принять", callback_data=f"mod_accept:{receipt['row']}"))
        kb.row(InlineKeyboardButton(
            text="⚡ Быстрое отклонение",
            callback_data=f"mod_reject_photo:{receipt['row']}",
        ))
        kb.row(InlineKeyboardButton(
            text="💬 Отклонить с комментарием",
            callback_data=f"mod_reject_custom:{receipt['row']}",
        ))
        reply_markup = kb.as_markup()
    elif with_history_buttons:
        kb = InlineKeyboardBuilder()
        kb.row(InlineKeyboardButton(text="🗑 Удалить чек", callback_data=f"hist_delete:{receipt['row']}"))
        date_str = receipt["date"][:10] if receipt.get("date") else ""
        if date_str:
            kb.row(InlineKeyboardButton(text="⬅️ Назад к списку", callback_data=f"cal_day:{date_str}"))
        reply_markup = kb.as_markup()

    if receipt["file_id"]:
        try:
            await message.answer_photo(receipt["file_id"], caption=caption, reply_markup=reply_markup)
            return
        except Exception:
            logger.exception("Не удалось отправить фото чека администратору")

    await message.answer(caption, reply_markup=reply_markup)


def _truncate(text: str, max_len: int) -> str:
    text = text or ""
    if len(text) > max_len:
        return text[: max_len - 1] + "…"
    return text


def _format_receipt_label(r: dict) -> str:
    """Короткая подпись для кнопки в списке чеков: имя (или ID), магазин,
    username. Каждая часть обрезается ОТДЕЛЬНО (а не всё целиком), иначе
    при длинном ФИО магазин/username могли полностью "вытесняться" за
    пределы общего лимита длины текста кнопки."""
    reg = google_sheets.get_registration_by_telegram_id(r["telegram_id"])
    name = _truncate(reg["full_name"] if reg else f"ID {r['telegram_id']}", 12)
    shop = _truncate(reg.get("shop") if reg else "", 10)
    username = _truncate(f"@{r['username']}" if r.get("username") else "", 10)

    parts = [p for p in [name, shop, username] if p]
    return " · ".join(parts)


# ---------- История: календарь ----------

MONTHS_RU = [
    "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]


def _build_calendar(year: int, month: int):
    builder = InlineKeyboardBuilder()

    builder.row(InlineKeyboardButton(text=f"{MONTHS_RU[month]} {year}", callback_data="cal_ignore"))
    builder.row(*[
        InlineKeyboardButton(text=d, callback_data="cal_ignore")
        for d in ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    ])

    for week in calendar.monthcalendar(year, month):
        row = []
        for day in week:
            if day == 0:
                row.append(InlineKeyboardButton(text=" ", callback_data="cal_ignore"))
            else:
                row.append(InlineKeyboardButton(
                    text=str(day),
                    callback_data=f"cal_day:{year:04d}-{month:02d}-{day:02d}",
                ))
        builder.row(*row)

    prev_year, prev_month = (year - 1, 12) if month == 1 else (year, month - 1)
    next_year, next_month = (year + 1, 1) if month == 12 else (year, month + 1)
    builder.row(
        InlineKeyboardButton(text="◀️", callback_data=f"cal_nav:{prev_year:04d}-{prev_month:02d}"),
        InlineKeyboardButton(text="▶️", callback_data=f"cal_nav:{next_year:04d}-{next_month:02d}"),
    )
    return builder.as_markup()


@dp.message(F.text == ADMIN_BTN_HISTORY, IsModerator())
async def admin_history_start(message: Message):
    today = google_sheets.moscow_today()
    await message.answer("Выберите дату:", reply_markup=_build_calendar(today.year, today.month))


@dp.callback_query(F.data == "cal_ignore", IsModerator())
async def admin_calendar_ignore(call: CallbackQuery):
    await call.answer()


@dp.callback_query(F.data.startswith("cal_nav:"), IsModerator())
async def admin_calendar_nav(call: CallbackQuery):
    _, ym = call.data.split(":", 1)
    year, month = map(int, ym.split("-"))
    await call.message.edit_reply_markup(reply_markup=_build_calendar(year, month))
    await call.answer()


async def _send_day_receipt_list(message: Message, date_str: str):
    """Показывает список чеков за дату — используется и при выборе даты в
    календаре, и как "Назад" из карточки чека, и после удаления чека."""
    try:
        receipts = google_sheets.get_receipts_by_date(date_str)
    except Exception:
        logger.exception("Не удалось получить чеки за дату")
        await message.answer("Не получилось прочитать данные из таблицы, попробуйте ещё раз.")
        return

    if not receipts:
        await message.answer(f"На {date_str} чеков нет.")
        return

    builder = InlineKeyboardBuilder()
    for r in receipts:
        label = _format_receipt_label(r)
        time_part = r["date"][11:16] if len(r["date"]) >= 16 else r["date"]
        builder.row(InlineKeyboardButton(
            text=f"{time_part} — {label} ({r['status']})",
            callback_data=f"hist_view:{r['row']}",
        ))
    await message.answer(f"Чеки за {date_str}:", reply_markup=builder.as_markup())


@dp.callback_query(F.data.startswith("cal_day:"), IsModerator())
async def admin_calendar_pick_day(call: CallbackQuery):
    _, date_str = call.data.split(":", 1)
    await _send_day_receipt_list(call.message, date_str)
    await call.answer()


@dp.callback_query(F.data.startswith("hist_view:"), IsModerator())
async def admin_history_view(call: CallbackQuery):
    row = int(call.data.split(":", 1)[1])
    receipt = google_sheets.get_receipt_by_row(row)
    if not receipt:
        await call.answer("Не найдено", show_alert=True)
        return
    await _send_receipt_card(call.message, receipt, with_history_buttons=True)
    await call.answer()


@dp.callback_query(F.data.startswith("hist_delete:"), IsModerator())
async def admin_history_delete_confirm(call: CallbackQuery):
    row = int(call.data.split(":", 1)[1])
    kb = InlineKeyboardBuilder()
    kb.row(
        InlineKeyboardButton(text="✅ Да, удалить", callback_data=f"hist_delete_yes:{row}"),
        InlineKeyboardButton(text="❌ Отмена", callback_data="hist_delete_no"),
    )
    await call.message.answer(
        "Точно удалить этот чек из таблицы? Действие необратимо.",
        reply_markup=kb.as_markup(),
    )
    await call.answer()


@dp.callback_query(F.data == "hist_delete_no", IsModerator())
async def admin_history_delete_cancel(call: CallbackQuery):
    await call.message.answer("Отменено, чек не удалён.")
    await call.answer()


@dp.callback_query(F.data.startswith("hist_delete_yes:"), IsModerator())
async def admin_history_delete_do(call: CallbackQuery):
    row = int(call.data.split(":", 1)[1])
    receipt = google_sheets.get_receipt_by_row(row)
    if not receipt:
        await call.answer("Уже удалено или не найдено.", show_alert=True)
        return

    date_str = receipt["date"][:10] if receipt.get("date") else ""

    try:
        google_sheets.mark_receipt_deleted(row)
    except Exception:
        logger.exception("Не удалось пометить чек как удалённый")
        await call.message.answer("Не получилось обновить таблицу, попробуйте ещё раз.")
        await call.answer()
        return

    await call.message.answer(
        "Чек помечен как удалённый и больше не будет виден в списках. "
        "Сама запись и фото остаются в таблице/на диске."
    )
    if date_str:
        await _send_day_receipt_list(call.message, date_str)
    await call.answer()


# ---------- Модерация чеков ----------

@dp.message(F.text == ADMIN_BTN_MODERATION, IsModerator())
async def admin_moderation_menu(message: Message):
    kb = InlineKeyboardBuilder()
    kb.row(InlineKeyboardButton(text="📋 Список на модерации", callback_data="moderation_list"))
    kb.row(InlineKeyboardButton(text="🤖 Автоматическая модерация", callback_data="moderation_auto"))
    await message.answer("Что открыть?", reply_markup=kb.as_markup())


async def _show_moderation_list(message: Message):
    try:
        receipts = google_sheets.get_pending_receipts()
    except Exception:
        logger.exception("Не удалось получить чеки на модерации")
        await message.answer("Не получилось прочитать данные из таблицы, попробуйте ещё раз.")
        return

    if not receipts:
        await message.answer("Чеков на модерации нет.")
        return

    builder = InlineKeyboardBuilder()
    for r in receipts:
        label = _format_receipt_label(r)
        builder.row(InlineKeyboardButton(text=f"{r['date']} — {label}", callback_data=f"mod_view:{r['row']}"))
    await message.answer("Чеки на модерации:", reply_markup=builder.as_markup())


@dp.callback_query(F.data == "moderation_list", IsModerator())
async def admin_moderation_list_cb(call: CallbackQuery):
    await call.answer()
    await _show_moderation_list(call.message)


@dp.callback_query(F.data.startswith("mod_view:"), IsModerator())
async def admin_moderation_view(call: CallbackQuery):
    row = int(call.data.split(":", 1)[1])
    receipt = google_sheets.get_receipt_by_row(row)
    if not receipt:
        await call.answer("Не найдено", show_alert=True)
        return
    await _send_receipt_card(call.message, receipt, with_moderation_buttons=True)
    await call.answer()


async def _reject_receipt(row: int, reason_text: str):
    """Ставит чеку статус 'отклонён' и уведомляет пользователя причиной.
    Возвращает (True, None) при успехе или (False, текст_ошибки)."""
    receipt = google_sheets.get_receipt_by_row(row)
    if not receipt:
        return False, "Не найдено."

    try:
        google_sheets.update_receipt_status(row, google_sheets.STATUS_REJECTED, comment=reason_text)
    except Exception:
        logger.exception("Не удалось обновить статус чека")
        return False, "Не получилось обновить таблицу, попробуйте ещё раз."

    try:
        await bot.send_message(receipt["telegram_id"], reason_text)
    except Exception:
        logger.exception("Не удалось уведомить пользователя об отклонении")

    return True, None


@dp.callback_query(F.data.startswith("mod_reject_photo:"), IsModerator())
async def admin_moderation_reject_photo(call: CallbackQuery):
    row = int(call.data.split(":", 1)[1])
    ok, error = await _reject_receipt(row, google_sheets.get_text("reject_message"))
    if not ok:
        await call.answer(error, show_alert=True)
        return
    await call.message.answer("Чек отклонён (быстрое отклонение), пользователь уведомлён.")
    await call.answer()


@dp.callback_query(F.data.startswith("mod_reject_custom:"), IsModerator())
async def admin_moderation_reject_custom_start(call: CallbackQuery, state: FSMContext):
    row = int(call.data.split(":", 1)[1])
    await state.set_state(AdminForm.waiting_reject_reason)
    await state.update_data(moderation_row=row)
    await call.message.answer(
        "Введите причину отклонения одним сообщением — она будет отправлена пользователю."
    )
    await call.answer()


@dp.message(StateFilter(AdminForm.waiting_reject_reason), IsModerator())
async def admin_moderation_reject_custom_finish(message: Message, state: FSMContext):
    data = await state.get_data()
    row = data.get("moderation_row")

    ok, error = await _reject_receipt(row, message.text)
    await state.clear()

    menu = _menu_for(message.from_user.id)
    if not ok:
        await message.answer(error, reply_markup=menu)
        return
    await message.answer("Чек отклонён, пользователь уведомлён.", reply_markup=menu)


@dp.callback_query(F.data.startswith("mod_accept:"), IsModerator())
async def admin_moderation_accept_start(call: CallbackQuery, state: FSMContext):
    row = int(call.data.split(":", 1)[1])
    await state.set_state(AdminForm.waiting_coupon)
    await state.update_data(moderation_row=row)
    await call.message.answer("Введите номер купона одним сообщением — я отправлю его пользователю.")
    await call.answer()


@dp.message(StateFilter(AdminForm.waiting_coupon), IsModerator())
async def admin_moderation_accept_finish(message: Message, state: FSMContext):
    data = await state.get_data()
    row = data.get("moderation_row")
    coupon = message.text

    receipt = google_sheets.get_receipt_by_row(row)
    if not receipt:
        await message.answer("Не удалось найти этот чек (возможно, таблица изменилась).")
        await state.clear()
        return

    try:
        google_sheets.update_receipt_status(row, google_sheets.STATUS_ACCEPTED, coupon=coupon)
    except Exception:
        logger.exception("Не удалось обновить статус чека")
        await message.answer("Не получилось обновить таблицу, попробуйте ещё раз.")
        await state.clear()
        return

    try:
        accept_text_template = google_sheets.get_text("accept_message")
        try:
            accept_text = accept_text_template.format(coupon=coupon)
        except Exception:
            logger.exception("Не удалось подставить купон в шаблон сообщения, использую текст по умолчанию")
            accept_text = f"Ваш чек принят! Ваш купон Ozon: {coupon}"

        await bot.send_message(receipt["telegram_id"], accept_text)
        await message.answer("Купон отправлен пользователю, статус обновлён.", reply_markup=_menu_for(message.from_user.id))
    except Exception:
        logger.exception("Не удалось отправить купон пользователю")
        await message.answer(
            "Статус обновлён, но отправить сообщение пользователю не удалось "
            "(возможно, он заблокировал бота).",
            reply_markup=_menu_for(message.from_user.id),
        )

    await state.clear()


# ---------- Управление модераторами (только владельцы из ADMIN_IDS) ----------

@dp.message(F.text == ADMIN_BTN_MODERATORS, IsOwner())
async def admin_moderators_menu(message: Message):
    try:
        moderators = google_sheets.get_moderators()
    except Exception:
        logger.exception("Не удалось получить список модераторов")
        await message.answer("Не получилось прочитать данные из таблицы, попробуйте ещё раз.")
        return

    builder = InlineKeyboardBuilder()
    for m in moderators:
        label = m["username"] or m["telegram_id"]
        builder.row(InlineKeyboardButton(text=f"❌ Удалить: {label}", callback_data=f"modrole_del:{m['row']}"))
    builder.row(InlineKeyboardButton(text="➕ Добавить модератора", callback_data="modrole_add"))

    if moderators:
        listing = "\n".join(
            f"- {m['telegram_id']}" + (f" (@{m['username']})" if m["username"] else "")
            for m in moderators
        )
    else:
        listing = "пока никого нет."

    await message.answer(f"Текущие модераторы:\n{listing}", reply_markup=builder.as_markup())


@dp.callback_query(F.data == "modrole_add", IsOwner())
async def admin_moderator_add_start(call: CallbackQuery, state: FSMContext):
    await state.set_state(AdminForm.waiting_moderator_id)
    await call.message.answer(
        "Пришлите Telegram ID нового модератора одним сообщением (число). "
        "Попросите его узнать свой ID через @userinfobot и прислать вам."
    )
    await call.answer()


@dp.message(StateFilter(AdminForm.waiting_moderator_id), IsOwner())
async def admin_moderator_add_finish(message: Message, state: FSMContext):
    text = (message.text or "").strip()
    if not text.isdigit():
        await message.answer("Это не похоже на Telegram ID — пришлите число, например 123456789.")
        return

    new_id = int(text)

    # Username нельзя узнать из одного только введённого числа — пробуем
    # спросить у Telegram напрямую. Это сработает, только если новый
    # модератор уже хотя бы раз писал этому боту раньше (иначе Telegram
    # не отдаёт информацию о чате) — если нет, username останется пустым,
    # это не критично: права проверяются по Telegram ID, а не по username.
    username = ""
    try:
        chat = await bot.get_chat(new_id)
        username = chat.username or ""
    except Exception:
        logger.info(
            "Не удалось получить username для %s (скорее всего, он ещё не писал боту) — "
            "сохраняю без username",
            new_id,
        )

    try:
        google_sheets.add_moderator(new_id, username=username, added_by=message.from_user.id)
    except Exception:
        logger.exception("Не удалось добавить модератора")
        await message.answer("Не получилось сохранить в таблицу, попробуйте ещё раз.")
        await state.clear()
        return

    await state.clear()
    await message.answer(f"Готово, {new_id} теперь модератор.", reply_markup=owner_menu)

    try:
        await bot.send_message(new_id, "Вас назначили модератором бота. Отправьте /admin, чтобы открыть панель.")
    except Exception:
        logger.exception("Не удалось уведомить нового модератора (возможно, он ещё не запускал бота)")


@dp.callback_query(F.data.startswith("modrole_del:"), IsOwner())
async def admin_moderator_delete(call: CallbackQuery):
    row = int(call.data.split(":", 1)[1])
    try:
        google_sheets.remove_moderator(row)
        await call.message.answer("Модератор удалён.")
    except Exception:
        logger.exception("Не удалось удалить модератора")
        await call.message.answer("Не получилось удалить, попробуйте ещё раз.")
    await call.answer()


# ---------- Редактирование текстов-автоответов (только владельцы) ----------

@dp.message(F.text == ADMIN_BTN_TEXTS, IsOwner())
async def admin_texts_menu(message: Message):
    builder = InlineKeyboardBuilder()
    for key, label in google_sheets.TEXT_LABELS.items():
        builder.row(InlineKeyboardButton(text=label, callback_data=f"text_edit:{key}"))
    await message.answer("Какой текст изменить?", reply_markup=builder.as_markup())


@dp.callback_query(F.data.startswith("text_edit:"), IsOwner())
async def admin_text_edit_start(call: CallbackQuery, state: FSMContext):
    key = call.data.split(":", 1)[1]
    label = google_sheets.TEXT_LABELS.get(key, key)

    try:
        current = google_sheets.get_text(key)
    except Exception:
        logger.exception("Не удалось прочитать текущий текст")
        await call.message.answer("Не получилось прочитать данные из таблицы, попробуйте ещё раз.")
        await call.answer()
        return

    await state.set_state(AdminForm.waiting_text_value)
    await state.update_data(text_key=key)
    await call.message.answer(
        f"«{label}»\n\nТекущий текст:\n{current}\n\nПришлите новый текст одним сообщением."
    )
    await call.answer()


@dp.message(StateFilter(AdminForm.waiting_text_value), IsOwner())
async def admin_text_edit_finish(message: Message, state: FSMContext):
    data = await state.get_data()
    key = data.get("text_key")
    new_value = message.text

    try:
        google_sheets.set_text(key, new_value)
    except Exception:
        logger.exception("Не удалось сохранить текст")
        await message.answer("Не получилось сохранить в таблицу, попробуйте ещё раз.")
        await state.clear()
        return

    await state.clear()
    await message.answer("Готово, текст обновлён.", reply_markup=owner_menu)


# ---------- Импорт (акционные позиции / отчёт по продажам) ----------

@dp.message(F.text == ADMIN_BTN_IMPORT, IsOwner())
async def admin_import_menu(message: Message):
    kb = InlineKeyboardBuilder()
    kb.row(InlineKeyboardButton(text="📦 Акционные позиции", callback_data="import_promo"))
    kb.row(InlineKeyboardButton(text="📈 Отчёт по продажам", callback_data="import_sales"))
    await message.answer(
        "Что загружаем? Оба варианта — файл .xlsx, первая строка листа "
        "должна быть заголовком (её содержимое не важно, бот её "
        "пропускает и читает данные со 2-й строки).\n\n"
        "📦 Акционные позиции — колонки: Артикул, Наименование. "
        "Новая загрузка полностью заменяет предыдущий список.\n"
        "📈 Отчёт по продажам — колонки: ИНН, Наименование, Артикул, "
        "Количество, Цена, № УПД, Дата продажи. Новая загрузка "
        "ДОБАВЛЯЕТСЯ к уже имеющимся данным (ничего не заменяется), и "
        "каждая строка дополнительно помечается датой и временем именно "
        "этой загрузки.",
        reply_markup=kb.as_markup(),
    )


@dp.callback_query(F.data == "import_promo", IsOwner())
async def admin_import_promo_start(call: CallbackQuery, state: FSMContext):
    await state.set_state(AdminForm.waiting_promo_file)
    await call.message.answer(
        "Пришлите файл .xlsx со списком акционных позиций "
        "(колонки: Артикул, Наименование)."
    )
    await call.answer()


@dp.callback_query(F.data == "import_sales", IsOwner())
async def admin_import_sales_start(call: CallbackQuery, state: FSMContext):
    await state.set_state(AdminForm.waiting_sales_file)
    await call.message.answer(
        "Пришлите файл .xlsx с отчётом по продажам (колонки: ИНН, "
        "Наименование, Артикул, Количество, Цена, № УПД, Дата продажи)."
    )
    await call.answer()


def _read_xlsx_rows(file_bytes: io.BytesIO, min_columns: int):
    """Читает .xlsx: первая строка — заголовок (пропускается), из
    остальных строк берёт первые min_columns колонок текстом. Полностью
    пустые строки пропускаются."""
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None or all(cell in (None, "") for cell in row):
            continue
        values = ["" if cell is None else str(cell).strip() for cell in row[:min_columns]]
        if len(values) < min_columns:
            values += [""] * (min_columns - len(values))
        rows.append(values)
    return rows


async def _download_document(message: Message) -> io.BytesIO:
    buf = io.BytesIO()
    await bot.download(message.document.file_id, destination=buf)
    buf.seek(0)
    return buf


@dp.message(StateFilter(AdminForm.waiting_promo_file), F.document)
async def admin_import_promo_finish(message: Message, state: FSMContext):
    if not message.document.file_name.lower().endswith((".xlsx", ".xlsm")):
        await message.answer("Нужен файл в формате .xlsx. Пришлите файл ещё раз.")
        return

    try:
        buf = await _download_document(message)
        rows = _read_xlsx_rows(buf, min_columns=2)
        google_sheets.import_promo_items(rows)
    except Exception:
        logger.exception("Не удалось импортировать акционные позиции")
        await message.answer(
            "Не получилось прочитать файл. Проверьте формат и попробуйте ещё раз.",
        )
        return

    await state.clear()
    await message.answer(f"Готово, загружено позиций: {len(rows)}.", reply_markup=owner_menu)


@dp.message(StateFilter(AdminForm.waiting_promo_file))
async def wrong_content_for_promo_file(message: Message):
    await message.answer("Пожалуйста, пришлите файл .xlsx (документом).")


@dp.message(StateFilter(AdminForm.waiting_sales_file), F.document)
async def admin_import_sales_finish(message: Message, state: FSMContext):
    if not message.document.file_name.lower().endswith((".xlsx", ".xlsm")):
        await message.answer("Нужен файл в формате .xlsx. Пришлите файл ещё раз.")
        return

    try:
        buf = await _download_document(message)
        rows = _read_xlsx_rows(buf, min_columns=7)
        google_sheets.import_sales_report(rows)
    except Exception:
        logger.exception("Не удалось импортировать отчёт по продажам")
        await message.answer(
            "Не получилось прочитать файл. Проверьте формат и попробуйте ещё раз.",
        )
        return

    await state.clear()
    await message.answer(
        f"Готово, добавлено строк: {len(rows)} (дата загрузки проставлена автоматически).",
        reply_markup=owner_menu,
    )


@dp.message(StateFilter(AdminForm.waiting_sales_file))
async def wrong_content_for_sales_file(message: Message):
    await message.answer("Пожалуйста, пришлите файл .xlsx (документом).")


# ---------- Отчёт в Excel ----------

def _autosize(ws, headers):
    for i, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(header) + 4)


def _build_report_workbook() -> io.BytesIO:
    registrations = google_sheets.get_all_registrations()
    reg_by_id = {reg["telegram_id"]: reg for reg in registrations}

    wb = openpyxl.Workbook()

    ws_receipts = wb.active
    ws_receipts.title = "Чеки"
    receipt_headers = [
        "Дата загрузки", "№ УПД", "Дата чека", "Дата регистрации", "Telegram ID", "Username",
        "ФИО", "Магазин", "ИНН магазина", "Телефон", "Статус", "Купон", "Комментарий", "Удалён",
    ]
    ws_receipts.append(receipt_headers)
    for r in google_sheets.get_receipts():
        reg = reg_by_id.get(r["telegram_id"])
        ws_receipts.append([
            r["date"],
            r["upd_number"],
            r["receipt_date"],
            reg["date"] if reg else "",
            r["telegram_id"],
            r["username"],
            reg["full_name"] if reg else "",
            reg["shop"] if reg else "",
            reg["inn"] if reg else "",
            reg["phone"] if reg else "",
            r["status"],
            r["coupon"],
            r["comment"],
            "да" if r["deleted"] else "",
        ])
    _autosize(ws_receipts, receipt_headers)

    ws_reg = wb.create_sheet("Регистрации")
    reg_headers = ["Дата регистрации", "ФИО", "Магазин", "ИНН магазина", "Телефон", "Telegram ID", "Username"]
    ws_reg.append(reg_headers)
    for reg in registrations:
        ws_reg.append([reg["date"], reg["full_name"], reg["shop"], reg["inn"], reg["phone"], reg["telegram_id"], reg["username"]])
    _autosize(ws_reg, reg_headers)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@dp.message(F.text == ADMIN_BTN_REPORT, IsModerator())
async def admin_reports_menu(message: Message):
    kb = InlineKeyboardBuilder()
    kb.row(InlineKeyboardButton(text="📄 Общий отчёт", callback_data="report_general"))
    kb.row(InlineKeyboardButton(text="🔎 Отчёт по продажам", callback_data="report_sales_match"))
    await message.answer("Какой отчёт сформировать?", reply_markup=kb.as_markup())


@dp.callback_query(F.data == "report_general", IsModerator())
async def admin_report_general(call: CallbackQuery):
    await call.answer()
    await call.message.answer("Формирую отчёт, секунду...")
    try:
        buf = _build_report_workbook()
    except Exception:
        logger.exception("Не удалось сформировать отчёт")
        await call.message.answer("Не получилось сформировать отчёт, попробуйте позже.")
        return

    filename = f"report_{google_sheets.moscow_today().strftime('%Y%m%d')}.xlsx"
    await call.message.answer_document(BufferedInputFile(buf.read(), filename=filename))


# ---------- Отчёт по продажам (сверка чеков с отчётом по продажам) ----------

def _parse_date_loose(text: str):
    """Разбирает дату в разных написаниях (точки/дефисы/слэши, 2 или 4
    цифры года) в объект date. Возвращает None, если не получилось."""
    if not text:
        return None
    normalized = re.sub(r"[/\-]", ".", text.strip())
    parts = normalized.split(".")
    if len(parts) != 3:
        return None
    day_s, month_s, year_s = parts
    if len(year_s) == 2:
        year_s = "20" + year_s
    try:
        return date(int(year_s), int(month_s), int(day_s))
    except ValueError:
        return None


_PERIOD_RE = re.compile(r"^(.+?)\s*-\s*(.+)$")


@dp.callback_query(F.data == "report_sales_match", IsModerator())
async def admin_report_sales_match_start(call: CallbackQuery, state: FSMContext):
    await state.set_state(AdminForm.waiting_sales_period)
    await call.message.answer(
        "Пришлите период для сверки в формате ДД.ММ.ГГГГ - ДД.ММ.ГГГГ, "
        "например 01.07.2026 - 22.07.2026.\n\n"
        "Бот сверит чеки с датой чека в этом периоде (у которых указаны "
        "ИНН, номер УПД и дата) с загруженным отчётом по продажам: "
        "совпадением считается одинаковые ИНН + № УПД + дата продажи, "
        "и хотя бы один товар по этому УПД — из списка акционных позиций."
    )
    await call.answer()


def _build_sales_match_workbook(start_date: date, end_date: date):
    """Сверяет чеки (ИНН магазина + № УПД + дата чека) с отчётом по
    продажам (ИНН + № УПД + дата продажи), проверяя, что хотя бы один
    товар по этому УПД входит в список акционных позиций. Возвращает
    (буфер с Excel, число совпадений-строк, число чеков без совпадения,
    список номеров строк чеков, у которых нашлось хотя бы одно
    совпадение — используется для массовой смены статуса)."""
    promo_articles = {
        item["article"].strip() for item in google_sheets.get_promo_items() if item["article"].strip()
    }

    sales_index = {}
    for s in google_sheets.get_sales_report():
        sale_date = _parse_date_loose(s["sale_date"])
        if not sale_date:
            continue
        key = (s["inn"].strip(), s["upd_number"].strip(), sale_date)
        sales_index.setdefault(key, []).append(s)

    matched = []
    unmatched = []

    for r in google_sheets.get_receipts():
        if r["deleted"] or not r["upd_number"] or not r["receipt_date"]:
            continue

        receipt_date = _parse_date_loose(r["receipt_date"])
        if not receipt_date or not (start_date <= receipt_date <= end_date):
            continue

        reg = google_sheets.get_registration_by_telegram_id(r["telegram_id"])
        inn = (reg.get("inn") or "").strip() if reg else ""

        if not inn:
            unmatched.append((r, reg, "у пользователя не указан ИНН магазина"))
            continue

        sale_lines = sales_index.get((inn, r["upd_number"].strip(), receipt_date), [])
        promo_lines = [s for s in sale_lines if s["article"].strip() in promo_articles]

        if promo_lines:
            for s in promo_lines:
                matched.append((r, reg, s))
        elif sale_lines:
            unmatched.append((r, reg, "товары по этому УПД не входят в акционные позиции"))
        else:
            unmatched.append((r, reg, "нет строки в отчёте по продажам с таким ИНН/№ УПД/датой"))

    wb = openpyxl.Workbook()

    ws_matched = wb.active
    ws_matched.title = "Совпадения"
    matched_headers = [
        "Дата чека", "№ УПД", "Telegram ID", "Username", "ФИО", "Магазин", "ИНН магазина",
        "Артикул", "Наименование товара", "Количество", "Цена",
    ]
    ws_matched.append(matched_headers)
    for r, reg, s in matched:
        ws_matched.append([
            r["receipt_date"], r["upd_number"], r["telegram_id"], r["username"],
            reg["full_name"] if reg else "", reg["shop"] if reg else "", reg["inn"] if reg else "",
            s["article"], s["name"], s["quantity"], s["price"],
        ])
    _autosize(ws_matched, matched_headers)

    ws_unmatched = wb.create_sheet("Без совпадений")
    unmatched_headers = [
        "Дата чека", "№ УПД", "Telegram ID", "Username", "ФИО", "Магазин", "ИНН магазина", "Причина",
    ]
    ws_unmatched.append(unmatched_headers)
    for r, reg, reason in unmatched:
        ws_unmatched.append([
            r["receipt_date"], r["upd_number"], r["telegram_id"], r["username"],
            reg["full_name"] if reg else "", reg["shop"] if reg else "", reg["inn"] if reg else "", reason,
        ])
    _autosize(ws_unmatched, unmatched_headers)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    matched_rows = sorted({r["row"] for r, _reg, _s in matched})
    return buf, len(matched), len(unmatched), matched_rows


@dp.message(StateFilter(AdminForm.waiting_sales_period), F.text)
async def admin_report_sales_match_finish(message: Message, state: FSMContext):
    m = _PERIOD_RE.match(message.text.strip())
    start_date = end_date = None
    if m:
        start_date = _parse_date_loose(m.group(1))
        end_date = _parse_date_loose(m.group(2))

    if not m or not start_date or not end_date or start_date > end_date:
        await message.answer(
            "Не получилось разобрать период. Введите в формате "
            "ДД.ММ.ГГГГ - ДД.ММ.ГГГГ, например 01.07.2026 - 22.07.2026."
        )
        return

    await message.answer("Сверяю данные, секунду...")

    try:
        buf, matched_count, unmatched_count, _matched_rows = _build_sales_match_workbook(start_date, end_date)
    except Exception:
        logger.exception("Не удалось выполнить сверку продаж")
        await message.answer("Не получилось выполнить сверку, попробуйте позже.")
        await state.clear()
        return

    await state.clear()
    filename = f"sverka_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
    await message.answer_document(
        BufferedInputFile(buf.read(), filename=filename),
        caption=f"Совпадений: {matched_count}. Без совпадений: {unmatched_count}.",
        reply_markup=_menu_for(message.from_user.id),
    )


@dp.message(StateFilter(AdminForm.waiting_sales_period))
async def wrong_content_for_sales_period(message: Message):
    await message.answer("Пожалуйста, пришлите период текстом, в формате ДД.ММ.ГГГГ - ДД.ММ.ГГГГ.")


# ---------- Автоматическая модерация (та же сверка + смена статуса) ----------

@dp.callback_query(F.data == "moderation_auto", IsModerator())
async def admin_moderation_auto_start(call: CallbackQuery, state: FSMContext):
    await state.set_state(AdminForm.waiting_auto_moderation_period)
    await call.message.answer(
        "Пришлите период для автоматической модерации в формате "
        "ДД.ММ.ГГГГ - ДД.ММ.ГГГГ, например 01.07.2026 - 22.07.2026.\n\n"
        "Бот сверит чеки за этот период с отчётом по продажам (так же, "
        "как «📊 Отчёты» -> «🔎 Отчёт по продажам»), пришлёт Excel и "
        "предложит сразу проставить статус «принят» найденным "
        "совпадениям."
    )
    await call.answer()


@dp.message(StateFilter(AdminForm.waiting_auto_moderation_period), F.text)
async def admin_moderation_auto_finish(message: Message, state: FSMContext):
    m = _PERIOD_RE.match(message.text.strip())
    start_date = end_date = None
    if m:
        start_date = _parse_date_loose(m.group(1))
        end_date = _parse_date_loose(m.group(2))

    if not m or not start_date or not end_date or start_date > end_date:
        await message.answer(
            "Не получилось разобрать период. Введите в формате "
            "ДД.ММ.ГГГГ - ДД.ММ.ГГГГ, например 01.07.2026 - 22.07.2026."
        )
        return

    await message.answer("Сверяю данные, секунду...")

    try:
        buf, matched_count, unmatched_count, matched_rows = _build_sales_match_workbook(start_date, end_date)
    except Exception:
        logger.exception("Не удалось выполнить автоматическую модерацию")
        await message.answer("Не получилось выполнить сверку, попробуйте позже.")
        await state.clear()
        return

    filename = f"avto_moderatsiya_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
    await message.answer_document(
        BufferedInputFile(buf.read(), filename=filename),
        caption=(
            f"Совпадений: {matched_count}. Без совпадений: {unmatched_count}. "
            f"Уникальных чеков с совпадением: {len(matched_rows)}."
        ),
    )

    if not matched_rows:
        await state.clear()
        await message.answer("Менять нечего — совпадений не найдено.", reply_markup=_menu_for(message.from_user.id))
        return

    await state.update_data(matched_rows=matched_rows)
    kb = InlineKeyboardBuilder()
    kb.row(
        InlineKeyboardButton(text="✅ Да", callback_data="auto_mod_apply_yes"),
        InlineKeyboardButton(text="❌ Нет", callback_data="auto_mod_apply_no"),
    )
    await message.answer(
        f"Изменить статус чеков на «принят» у {len(matched_rows)} шт. "
        "(совпавших по ИНН/№ УПД/дате продажи с акционным товаром)? "
        "Пользователям сообщение отправлено НЕ будет — купон в этом "
        "сценарии не назначается, меняется только статус в таблице.",
        reply_markup=kb.as_markup(),
    )


@dp.callback_query(F.data == "auto_mod_apply_yes", IsModerator())
async def admin_moderation_auto_apply_yes(call: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    rows = data.get("matched_rows", [])
    await state.clear()
    await call.answer()

    updated = 0
    for row in rows:
        try:
            google_sheets.update_receipt_status(
                row,
                google_sheets.STATUS_ACCEPTED,
                comment="Подтверждено автоматической сверкой с отчётом по продажам",
            )
            updated += 1
        except Exception:
            logger.exception("Не удалось обновить статус чека %s при автомодерации", row)

    await call.message.answer(f"Готово, статус изменён у {updated} из {len(rows)} чеков.")


@dp.callback_query(F.data == "auto_mod_apply_no", IsModerator())
async def admin_moderation_auto_apply_no(call: CallbackQuery, state: FSMContext):
    await state.clear()
    await call.answer()
    await call.message.answer("Хорошо, статусы не менял.")


@dp.message(StateFilter(AdminForm.waiting_auto_moderation_period))
async def wrong_content_for_auto_moderation_period(message: Message):
    await message.answer("Пожалуйста, пришлите период текстом, в формате ДД.ММ.ГГГГ - ДД.ММ.ГГГГ.")


async def main():
    # На Amvera (и вообще при передеплое на любом хостинге) старый процесс
    # бота может ещё секунду-другую держать соединение с Telegram после
    # того, как запустился новый. Без небольшой паузы это иногда даёт
    # ошибку "Terminated by other getUpdates request".
    startup_delay = float(os.getenv("STARTUP_DELAY_SECONDS", "3"))
    if startup_delay > 0:
        await asyncio.sleep(startup_delay)

    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
