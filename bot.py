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
    phone = State()
    waiting_photo = State()


class AdminForm(StatesGroup):
    waiting_coupon = State()
    waiting_moderator_id = State()
    waiting_text_value = State()
    waiting_reject_reason = State()


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
BTN_RULES = "❓ Правила"

ADMIN_BTN_HISTORY = "🗓 История"
ADMIN_BTN_MODERATION = "🛠 Модерация"
ADMIN_BTN_MODERATORS = "👥 Модераторы"
ADMIN_BTN_REPORT = "📊 Отчёт"
ADMIN_BTN_TEXTS = "✏️ Тексты"
ADMIN_BTN_EXIT = "⬅️ Выйти из панели"

_main_menu_builder = ReplyKeyboardBuilder()
_main_menu_builder.button(text=BTN_SEND_RECEIPT)
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
_owner_menu_builder.button(text=ADMIN_BTN_EXIT)
_owner_menu_builder.adjust(2, 2, 2)
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

    kb_builder = ReplyKeyboardBuilder()
    kb_builder.button(text="Поделиться контактом", request_contact=True)
    keyboard = kb_builder.as_markup(resize_keyboard=True, one_time_keyboard=True)

    await state.set_state(Form.phone)
    await message.answer(
        "И последнее: поделись номером телефона для связи.",
        reply_markup=keyboard,
    )


@dp.message(StateFilter(Form.phone), F.contact)
async def process_phone(message: Message, state: FSMContext):
    data = await state.get_data()
    full_name = data.get("name", "")
    shop = data.get("shop", "")
    phone = message.contact.phone_number

    try:
        google_sheets.append_registration(
            full_name=full_name,
            shop=shop,
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

    try:
        google_sheets.append_receipt(
            telegram_id=message.from_user.id,
            username=message.from_user.username,
            file_id=file_id,
            file_name=filepath,
        )
    except Exception:
        logger.exception("Не удалось записать чек в Google Таблицу")

    await state.clear()
    await message.answer(
        google_sheets.get_text("receipt_received"),
        reply_markup=main_menu,
    )


@dp.message(StateFilter(Form.waiting_photo))
async def wrong_content_for_photo(message: Message):
    await message.answer("Пожалуйста, пришлите именно фото (как изображение, не файлом).")


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
        lines.append(f"Телефон: {reg['phone']}")
    else:
        lines.append("Регистрация не найдена (данные могли не сохраниться).")

    lines.append(f"Telegram ID: {receipt['telegram_id']}")
    if receipt["username"]:
        lines.append(f"Username: @{receipt['username']}")
    lines.append(f"Дата: {receipt['date']}")
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
    today = date.today()
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
async def admin_moderation_start(message: Message):
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
        "Дата", "Telegram ID", "Username", "ФИО", "Магазин", "Телефон",
        "Статус", "Купон", "Комментарий", "Удалён",
    ]
    ws_receipts.append(receipt_headers)
    for r in google_sheets.get_receipts():
        reg = reg_by_id.get(r["telegram_id"])
        ws_receipts.append([
            r["date"],
            r["telegram_id"],
            r["username"],
            reg["full_name"] if reg else "",
            reg["shop"] if reg else "",
            reg["phone"] if reg else "",
            r["status"],
            r["coupon"],
            r["comment"],
            "да" if r["deleted"] else "",
        ])
    _autosize(ws_receipts, receipt_headers)

    ws_reg = wb.create_sheet("Регистрации")
    reg_headers = ["Дата регистрации", "ФИО", "Магазин", "Телефон", "Telegram ID", "Username"]
    ws_reg.append(reg_headers)
    for reg in registrations:
        ws_reg.append([reg["date"], reg["full_name"], reg["shop"], reg["phone"], reg["telegram_id"], reg["username"]])
    _autosize(ws_reg, reg_headers)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@dp.message(F.text == ADMIN_BTN_REPORT, IsModerator())
async def admin_report(message: Message):
    await message.answer("Формирую отчёт, секунду...")
    try:
        buf = _build_report_workbook()
    except Exception:
        logger.exception("Не удалось сформировать отчёт")
        await message.answer("Не получилось сформировать отчёт, попробуйте позже.")
        return

    filename = f"report_{date.today().strftime('%Y%m%d')}.xlsx"
    await message.answer_document(BufferedInputFile(buf.read(), filename=filename))


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
