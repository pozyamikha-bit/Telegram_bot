"""
Бот регистрации участников акции + приём чеков/УПД.

Сценарий:
  /start -> ФИО -> Магазин -> Телефон (кнопкой) -> запись в Google Таблицу
  -> главное меню -> "Отправить чек" -> фото -> запись в Google Таблицу

Перед запуском:
  1. Установите библиотеки:  pip install -r requirements.txt
  2. Скопируйте .env.example в .env и заполните значения
  3. Положите json-ключ сервисного аккаунта Google рядом
     (см. README.md, раздел "Настройка Google Sheets")
"""

import logging
import os

from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.utils import executor

import config
import google_sheets

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)

bot = Bot(token=config.BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

os.makedirs(config.RECEIPTS_DIR, exist_ok=True)


class Form(StatesGroup):
    name = State()
    shop = State()
    phone = State()
    waiting_photo = State()


BTN_SEND_RECEIPT = "📥 Отправить чек / УПД"
BTN_RULES = "❓ Правила"

main_menu = types.ReplyKeyboardMarkup(resize_keyboard=True)
main_menu.add(BTN_SEND_RECEIPT, BTN_RULES)


# ---------- Регистрация ----------

@dp.message_handler(commands=["start"])
async def cmd_start(message: types.Message, state: FSMContext):
    await state.finish()
    await Form.name.set()
    await message.reply(
        "Здравствуйте! Давайте зарегистрируем вас в программе.\n\n"
        "Для начала напишите, пожалуйста, ваше ФИО."
    )


@dp.message_handler(state=Form.name)
async def process_name(message: types.Message, state: FSMContext):
    await state.update_data(name=message.text)
    await Form.next()
    await message.reply("Отлично! Теперь укажите название вашего магазина.")


@dp.message_handler(state=Form.shop)
async def process_shop(message: types.Message, state: FSMContext):
    await state.update_data(shop=message.text)

    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    keyboard.add(types.KeyboardButton(text="Поделиться контактом", request_contact=True))

    await Form.next()
    await message.reply(
        "И последнее: поделись номером телефона для связи.",
        reply_markup=keyboard,
    )


@dp.message_handler(content_types=types.ContentType.CONTACT, state=Form.phone)
async def process_phone(message: types.Message, state: FSMContext):
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
        await message.reply(
            "Регистрация прошла, но возникла ошибка при записи в таблицу. "
            "Сообщите об этом администратору."
        )

    await state.finish()
    await message.reply(
        "Регистрация успешна! Теперь ты можешь отправлять чеки.",
        reply_markup=main_menu,
    )


# ---------- Правила ----------

@dp.message_handler(lambda message: message.text == BTN_RULES, state="*")
async def show_rules(message: types.Message):
    await message.reply(
        "Правила акции:\n"
        "1. Чек должен быть не старше 14 дней.\n"
        "2. На фото должны быть видны дата и название товара.\n"
        "3. Один чек — один купон Ozon.\n\n"
        "(Отредактируйте этот текст под условия своей акции.)"
    )


# ---------- Приём чеков ----------

@dp.message_handler(lambda message: message.text == BTN_SEND_RECEIPT, state="*")
async def send_receipt_start(message: types.Message):
    await Form.waiting_photo.set()
    await message.reply(
        "Пришлите ОДНО фото чека или УПД. "
        "На фото должно быть видно дату и название товара."
    )


@dp.message_handler(content_types=types.ContentType.PHOTO, state=Form.waiting_photo)
async def process_photo(message: types.Message, state: FSMContext):
    photo = message.photo[-1]  # берём фото в максимальном качестве
    file_id = photo.file_id

    filename = f"{message.from_user.id}_{message.date.strftime('%Y%m%d_%H%M%S')}.jpg"
    filepath = os.path.join(config.RECEIPTS_DIR, filename)

    try:
        await photo.download(destination_file=filepath)
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

    await state.finish()
    await message.reply(
        "Спасибо! Чек принят на модерацию. Ожидайте купон Ozon.",
        reply_markup=main_menu,
    )


@dp.message_handler(state=Form.waiting_photo, content_types=types.ContentType.ANY)
async def wrong_content_for_photo(message: types.Message):
    await message.reply("Пожалуйста, пришлите именно фото (как изображение, не файлом).")


if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)
