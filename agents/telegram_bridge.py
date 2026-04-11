"""
Telegram-бот для приёма заметок и фото в папку Объекты/Inbox/.

Запуск: uv run agents/telegram_bridge.py
Зависимости: aiogram, python-dotenv

Логика:
- Текст → дописывается в Объекты/Inbox/notes.md с меткой времени
- Фото → сохраняется в Объекты/Inbox/photos/<YYYY-MM-DD_HH-MM-SS>.jpg
- Все остальные пользователи игнорируются (их ID выводится в терминал)
"""

import asyncio
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.types import Message
from dotenv import load_dotenv

# --- Конфигурация ---

BASE_DIR = Path(__file__).parent.parent
INBOX_DIR = BASE_DIR / "Объекты" / "Inbox"
PHOTOS_DIR = INBOX_DIR / "photos"
NOTES_FILE = INBOX_DIR / "notes.md"

load_dotenv(BASE_DIR / ".env")

TOKEN = os.getenv("TELEGRAM_TOKEN")
if not TOKEN:
    sys.exit("Ошибка: TELEGRAM_TOKEN не найден в .env")

# Если ALLOWED_USER_ID не задан — бот работает в режиме обнаружения ID
_raw_id = os.getenv("ALLOWED_USER_ID", "").strip()
ALLOWED_USER_ID: int | None = int(_raw_id) if _raw_id.isdigit() else None

# --- Логирование ---

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# --- Инициализация ---

INBOX_DIR.mkdir(parents=True, exist_ok=True)
PHOTOS_DIR.mkdir(parents=True, exist_ok=True)

bot = Bot(token=TOKEN)
dp = Dispatcher()


# --- Хелперы ---

def _now_str(fmt: str = "%Y-%m-%d %H:%M:%S") -> str:
    return datetime.now().strftime(fmt)


def _is_allowed(user_id: int) -> bool:
    if ALLOWED_USER_ID is None:
        return True  # режим обнаружения — принимаем всех, но сообщаем ID
    return user_id == ALLOWED_USER_ID


def _append_note(text: str, source: str = "text") -> None:
    """Дописывает заметку в notes.md."""
    entry = f"\n## {_now_str()} [{source}]\n\n{text}\n"
    with open(NOTES_FILE, "a", encoding="utf-8") as f:
        if NOTES_FILE.stat().st_size == 0 if NOTES_FILE.exists() else True:
            f.write("# Inbox — заметки\n")
        f.write(entry)


# --- Хендлеры ---

@dp.message(CommandStart())
async def cmd_start(message: Message) -> None:
    uid = message.from_user.id
    if ALLOWED_USER_ID is None:
        log.info(f"[START] User ID: {uid} | @{message.from_user.username}")
        print(f"\n>>> Твой Telegram ID: {uid}")
        print(f">>> Запиши его в .env: ALLOWED_USER_ID={uid}\n")
        await message.answer(
            f"Бот запущен в режиме обнаружения ID.\n"
            f"Твой ID: `{uid}`\n\n"
            f"Запиши его в `.env`:\n`ALLOWED_USER_ID={uid}`",
            parse_mode="Markdown",
        )
    elif _is_allowed(uid):
        await message.answer("Готов! Отправляй текст или фото — всё попадёт в Inbox.")
    else:
        log.warning(f"[BLOCKED] Попытка доступа: user_id={uid}")


@dp.message(F.text)
async def handle_text(message: Message) -> None:
    uid = message.from_user.id

    if not _is_allowed(uid):
        log.warning(f"[BLOCKED text] user_id={uid}")
        return

    if ALLOWED_USER_ID is None:
        print(f"\n>>> Твой Telegram ID: {uid}")
        print(f">>> Запиши его в .env: ALLOWED_USER_ID={uid}\n")

    text = message.text.strip()
    _append_note(text, source="text")
    log.info(f"[NOTE] Сохранено: {len(text)} символов")
    await message.answer("✓ Заметка сохранена")


@dp.message(F.photo)
async def handle_photo(message: Message) -> None:
    uid = message.from_user.id

    if not _is_allowed(uid):
        log.warning(f"[BLOCKED photo] user_id={uid}")
        return

    if ALLOWED_USER_ID is None:
        print(f"\n>>> Твой Telegram ID: {uid}")
        print(f">>> Запиши его в .env: ALLOWED_USER_ID={uid}\n")

    # Берём самое большое фото
    photo = message.photo[-1]
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{timestamp}_{message.message_id}.jpg"
    save_path = PHOTOS_DIR / filename

    # Скачиваем файл
    file = await bot.get_file(photo.file_id)
    await bot.download_file(file.file_path, destination=save_path)

    # Если есть подпись — сохраняем её как заметку
    if message.caption:
        _append_note(f"📷 {filename}\n\n{message.caption.strip()}", source="photo")

    log.info(f"[PHOTO] Сохранено: {save_path}")
    await message.answer(f"✓ Фото сохранено: `{filename}`", parse_mode="Markdown")


@dp.message(F.document)
async def handle_document(message: Message) -> None:
    uid = message.from_user.id

    if not _is_allowed(uid):
        log.warning(f"[BLOCKED document] user_id={uid}")
        return

    doc = message.document
    mime = doc.mime_type or ""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    orig_name = doc.file_name or "file"

    if mime.startswith("image/"):
        # Изображение, отправленное без сжатия → в photos/
        save_dir = PHOTOS_DIR
        filename = f"{timestamp}_{message.message_id}_{orig_name}"
        category = "photo-doc"
    else:
        # Любой другой документ (PDF, DOCX, XLSX…) → в files/
        files_dir = INBOX_DIR / "files"
        files_dir.mkdir(parents=True, exist_ok=True)
        save_dir = files_dir
        filename = f"{timestamp}_{message.message_id}_{orig_name}"
        category = "file"

    save_path = save_dir / filename

    file = await bot.get_file(doc.file_id)
    await bot.download_file(file.file_path, destination=save_path)

    if message.caption:
        _append_note(f"📎 {filename}\n\n{message.caption.strip()}", source=category)

    log.info(f"[{category.upper()}] Сохранено: {save_path}")
    await message.answer(f"✓ Файл сохранён: `{filename}`", parse_mode="Markdown")


# --- Запуск ---

async def main() -> None:
    if ALLOWED_USER_ID is None:
        log.info("Режим обнаружения ID: ALLOWED_USER_ID не задан в .env")
        log.info("Напиши боту /start — в терминале появится твой ID")
    else:
        log.info(f"Бот запущен. Разрешённый user_id: {ALLOWED_USER_ID}")

    log.info(f"Inbox: {INBOX_DIR.resolve()}")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
