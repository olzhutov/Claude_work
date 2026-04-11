"""
Telegram-бот для приёма заметок, фото и документов в Объекты/.

Запуск: uv run agents/telegram_bridge.py
Зависимости: aiogram, python-dotenv

Логика:
- /menu → inline-клавиатура выбора папки проекта
- /run_extractor → запускает agents/extractor.py для текущей папки
- Текст → notes.md (Inbox или корень проекта)
- Фото (сжатые) → photos/ (Inbox) или корень проекта
- Документы → photos/ или files/ (Inbox) / корень проекта
- Фильтр по ALLOWED_USER_ID
"""

import asyncio
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    BotCommand,
    BufferedInputFile,
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from dotenv import load_dotenv

# --- Конфигурация ---

BASE_DIR = Path(__file__).parent.parent
OBJECTS_DIR = BASE_DIR / "Объекты"
INBOX_DIR = OBJECTS_DIR / "Inbox"
PHOTOS_DIR = INBOX_DIR / "photos"
NOTES_FILE = INBOX_DIR / "notes.md"

# Папки в корне Объекты/, которые не являются активными проектами
SKIP_DIRS = {"Inbox", "Архив", "__pycache__"}

load_dotenv(BASE_DIR / ".env")

TOKEN = os.getenv("TELEGRAM_TOKEN")
if not TOKEN:
    sys.exit("Ошибка: TELEGRAM_TOKEN не найден в .env")

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

# Текущий режим для каждого пользователя: 'Inbox' или имя папки объекта
active_sessions: dict[int, str] = {}


# --- Хелперы ---

def _now_str(fmt: str = "%Y-%m-%d %H:%M:%S") -> str:
    return datetime.now().strftime(fmt)


def _is_allowed(user_id: int) -> bool:
    if ALLOWED_USER_ID is None:
        return True
    return user_id == ALLOWED_USER_ID


def _get_session(user_id: int) -> str:
    return active_sessions.get(user_id, "Inbox")


def _scan_projects() -> list[str]:
    """Возвращает активные проекты из корня Объекты/ (только прямые подпапки)."""
    projects = []
    for p in sorted(OBJECTS_DIR.iterdir()):
        if (
            p.is_dir()
            and not p.name.startswith(".")
            and p.name not in SKIP_DIRS
        ):
            projects.append(p.name)
    return projects


def _build_menu_keyboard() -> InlineKeyboardMarkup:
    """Строит inline-клавиатуру со списком проектов + Inbox."""
    projects = _scan_projects()
    buttons = [
        [InlineKeyboardButton(text=name, callback_data=f"dest:{name}")]
        for name in projects
    ]
    buttons.append([
        InlineKeyboardButton(text="📥 Сбросить в Inbox", callback_data="dest:Inbox")
    ])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def _build_session_keyboard(session: str) -> InlineKeyboardMarkup | None:
    """Возвращает клавиатуру действий для выбранного проекта (не Inbox)."""
    if session == "Inbox":
        return None
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="📄 Получить сводку", callback_data=f"summary:{session}"),
    ]])


BINARY_EXTS = {".pdf", ".doc", ".docx", ".xls", ".xlsx"}
TEXT_EXTS   = {".md", ".txt"}
TG_MSG_LIMIT = 4096


def _collect_summary_items(project_name: str) -> tuple[list[Path], list[Path]]:
    """
    Собирает файлы для сводки:
      binary_docs — .pdf/.doc/.docx/.xls/.xlsx из корня папки проекта
      text_files  — wiki/info_brief.md, затем wiki/objects/*.md (или *.md из корня)
    Возвращает (binary_docs, text_files).
    """
    project_dir = OBJECTS_DIR / project_name

    # --- Бинарные документы из корня ---
    binary_docs = sorted(
        p for p in project_dir.iterdir()
        if p.is_file() and p.suffix.lower() in BINARY_EXTS
    )

    # --- Текстовые сводки ---
    text_files: list[Path] = []

    # 1. wiki/info_brief.md — главный приоритет
    info_brief = project_dir / "wiki" / "info_brief.md"
    if info_brief.exists():
        text_files.append(info_brief)
    else:
        # 2. wiki/objects/*.md
        objects_dir = project_dir / "wiki" / "objects"
        if objects_dir.is_dir():
            text_files.extend(sorted(objects_dir.glob("*.md")))

        # 3. *.md / *.txt в корне (кроме notes.md)
        if not text_files:
            for p in sorted(project_dir.iterdir()):
                if p.is_file() and p.suffix.lower() in TEXT_EXTS and p.name != "notes.md":
                    text_files.append(p)

    return binary_docs, text_files


def _split_text(text: str, limit: int = TG_MSG_LIMIT) -> list[str]:
    """Разбивает длинный текст на части, не превышающие лимит Telegram."""
    parts = []
    while len(text) > limit:
        # Ищем последний перенос строки в пределах лимита
        split_at = text.rfind("\n", 0, limit)
        if split_at == -1:
            split_at = limit
        parts.append(text[:split_at])
        text = text[split_at:].lstrip("\n")
    if text:
        parts.append(text)
    return parts


def _resolve_save_paths(session: str, filename: str, is_image: bool) -> Path:
    if session == "Inbox":
        if is_image:
            return PHOTOS_DIR / filename
        else:
            files_dir = INBOX_DIR / "files"
            files_dir.mkdir(parents=True, exist_ok=True)
            return files_dir / filename
    else:
        project_dir = OBJECTS_DIR / session
        project_dir.mkdir(parents=True, exist_ok=True)
        return project_dir / filename


def _append_note(text: str, source: str = "text", session: str = "Inbox") -> None:
    if session == "Inbox":
        notes_path = NOTES_FILE
        header = "# Inbox — заметки\n"
    else:
        notes_path = OBJECTS_DIR / session / "notes.md"
        header = f"# Заметки — {session}\n"

    entry = f"\n## {_now_str()} [{source}]\n\n{text}\n"
    is_new = not notes_path.exists() or notes_path.stat().st_size == 0
    with open(notes_path, "a", encoding="utf-8") as f:
        if is_new:
            f.write(header)
        f.write(entry)


def _session_label(session: str) -> str:
    return "📥 Inbox" if session == "Inbox" else f"📁 {session}"


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
        session = _get_session(uid)
        await message.answer(
            f"Готов! Текущая папка: {_session_label(session)}\n"
            f"Используй /menu для смены объекта.",
            reply_markup=_build_session_keyboard(session),
        )
    else:
        log.warning(f"[BLOCKED] Попытка доступа: user_id={uid}")


@dp.message(Command("menu"))
async def cmd_menu(message: Message) -> None:
    uid = message.from_user.id
    if not _is_allowed(uid):
        return
    session = _get_session(uid)
    await message.answer(
        f"Текущая папка: {_session_label(session)}\n\nВыберите объект для сохранения:",
        reply_markup=_build_menu_keyboard(),
    )


@dp.message(Command("run_extractor"))
async def cmd_run_extractor(message: Message) -> None:
    uid = message.from_user.id
    if not _is_allowed(uid):
        return

    session = _get_session(uid)
    if session == "Inbox":
        await message.answer("⚠️ Выберите конкретный объект через /menu перед запуском экстрактора.")
        return

    folder_path = OBJECTS_DIR / session
    await message.answer(f"⏳ Запускаю экстрактор для `{session}`...", parse_mode="Markdown")
    log.info(f"[EXTRACTOR] Запуск для {folder_path}")

    try:
        proc = await asyncio.create_subprocess_exec(
            sys.executable,
            str(BASE_DIR / "agents" / "extractor.py"),
            str(folder_path),
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            cwd=str(BASE_DIR),
        )
        stdout, stderr = await proc.communicate()

        if proc.returncode == 0:
            log.info(f"[EXTRACTOR] Завершён успешно для {session}")
            await message.answer(f"✅ Обработка завершена для `{session}`.", parse_mode="Markdown")
        else:
            err = stderr.decode("utf-8", errors="replace")[-800:]
            log.error(f"[EXTRACTOR] Ошибка: {err}")
            await message.answer(
                f"❌ Ошибка экстрактора:\n```\n{err}\n```",
                parse_mode="Markdown",
            )
    except Exception as e:
        log.exception("[EXTRACTOR] Исключение")
        await message.answer(f"❌ Не удалось запустить экстрактор: {e}")


@dp.callback_query(F.data.startswith("dest:"))
async def cb_set_destination(callback: CallbackQuery) -> None:
    uid = callback.from_user.id
    if not _is_allowed(uid):
        await callback.answer("Доступ запрещён.", show_alert=True)
        return

    dest = callback.data.removeprefix("dest:")
    active_sessions[uid] = dest
    label = _session_label(dest)

    await callback.message.edit_text(
        f"✅ Режим изменён. Теперь все файлы сохраняются в папку: {label}",
        reply_markup=_build_session_keyboard(dest),
    )
    await callback.answer()
    log.info(f"[SESSION] user={uid} → {dest}")


@dp.callback_query(F.data.startswith("summary:"))
async def cb_get_summary(callback: CallbackQuery) -> None:
    uid = callback.from_user.id
    if not _is_allowed(uid):
        await callback.answer("Доступ запрещён.", show_alert=True)
        return

    project_name = callback.data.removeprefix("summary:")
    await callback.answer()

    binary_docs, text_files = _collect_summary_items(project_name)

    if not binary_docs and not text_files:
        await callback.message.answer("Документов или справок пока нет.")
        log.info(f"[SUMMARY] Ничего не найдено для {project_name}")
        return

    # --- Отправка бинарных документов ---
    for doc_path in binary_docs:
        log.info(f"[SUMMARY] Документ → {doc_path.name}")
        await callback.message.answer_document(
            FSInputFile(doc_path, filename=doc_path.name),
            caption=f"📎 {doc_path.name}",
        )

    # --- Отправка текстовых сводок ---
    for txt_path in text_files:
        log.info(f"[SUMMARY] Текст → {txt_path.name}")
        content = txt_path.read_text(encoding="utf-8").strip()
        if not content:
            continue
        header = f"📄 *{txt_path.name}*\n\n"
        parts = _split_text(header + content)
        for i, part in enumerate(parts):
            await callback.message.answer(
                part,
                parse_mode="Markdown" if i == 0 else None,
            )


@dp.message(F.text)
async def handle_text(message: Message) -> None:
    uid = message.from_user.id
    if not _is_allowed(uid):
        log.warning(f"[BLOCKED text] user_id={uid}")
        return
    if ALLOWED_USER_ID is None:
        print(f"\n>>> Твой Telegram ID: {uid}\n>>> ALLOWED_USER_ID={uid}\n")

    session = _get_session(uid)
    text = message.text.strip()
    _append_note(text, source="text", session=session)
    log.info(f"[NOTE] {len(text)} символов → {session}")
    await message.answer(f"✓ Заметка сохранена ({_session_label(session)})")


@dp.message(F.photo)
async def handle_photo(message: Message) -> None:
    uid = message.from_user.id
    if not _is_allowed(uid):
        log.warning(f"[BLOCKED photo] user_id={uid}")
        return
    if ALLOWED_USER_ID is None:
        print(f"\n>>> Твой Telegram ID: {uid}\n>>> ALLOWED_USER_ID={uid}\n")

    session = _get_session(uid)
    photo = message.photo[-1]
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{timestamp}_{message.message_id}.jpg"

    save_path = _resolve_save_paths(session, filename, is_image=True)
    file = await bot.get_file(photo.file_id)
    await bot.download_file(file.file_path, destination=save_path)

    if message.caption:
        _append_note(f"📷 {filename}\n\n{message.caption.strip()}", source="photo", session=session)

    log.info(f"[PHOTO] {save_path}")
    await message.answer(
        f"✓ Фото сохранено: `{filename}`\n📁 {_session_label(session)}",
        parse_mode="Markdown",
    )


@dp.message(F.document)
async def handle_document(message: Message) -> None:
    uid = message.from_user.id
    if not _is_allowed(uid):
        log.warning(f"[BLOCKED document] user_id={uid}")
        return

    session = _get_session(uid)
    doc = message.document
    mime = doc.mime_type or ""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    orig_name = doc.file_name or "file"
    filename = f"{timestamp}_{message.message_id}_{orig_name}"
    is_image = mime.startswith("image/")
    category = "photo-doc" if is_image else "file"

    save_path = _resolve_save_paths(session, filename, is_image=is_image)
    file = await bot.get_file(doc.file_id)
    await bot.download_file(file.file_path, destination=save_path)

    if message.caption:
        _append_note(f"📎 {filename}\n\n{message.caption.strip()}", source=category, session=session)

    log.info(f"[{category.upper()}] {save_path}")
    await message.answer(
        f"✓ Файл сохранён: `{filename}`\n📁 {_session_label(session)}",
        parse_mode="Markdown",
    )


# --- Запуск ---

async def main() -> None:
    await bot.set_my_commands([
        BotCommand(command="start", description="Статус бота"),
        BotCommand(command="menu", description="Выбрать папку объекта"),
        BotCommand(command="run_extractor", description="Запустить экстрактор для текущего объекта"),
    ])

    if ALLOWED_USER_ID is None:
        log.info("Режим обнаружения ID: напиши /start")
    else:
        log.info(f"Бот запущен. Разрешённый user_id: {ALLOWED_USER_ID}")

    log.info(f"Объекты: {OBJECTS_DIR.resolve()}")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
