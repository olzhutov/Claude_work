"""
Анализатор клипингов недвижимости.

Сканирует Clippings/*.md, для каждого файла без `parsed: true` в YAML-frontmatter:
  1. Читает текст заметки и находит вложенные изображения (опционально)
  2. Отправляет в Claude API (text + vision) для извлечения структурированных данных
  3. Дописывает извлечённые поля и `parsed: true` обратно в YAML-frontmatter

Запуск:
    uv run agents/clip_analyzer.py                  # все необработанные файлы
    uv run agents/clip_analyzer.py --dry-run        # без записи, только вывод JSON
    uv run agents/clip_analyzer.py --reparse        # переобработать даже parsed: true
    uv run agents/clip_analyzer.py --file "path.md" # один конкретный файл

Зависимости: anthropic, PyYAML, python-dotenv (все уже в проекте)
"""

import argparse
import base64
import json
import logging
import os
import re
import sys
import urllib.request
from pathlib import Path

import yaml
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Конфигурация
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent.parent
CLIPPINGS_DIR = BASE_DIR / "Clippings"

load_dotenv(BASE_DIR / ".env")

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
if not ANTHROPIC_API_KEY:
    sys.exit(
        "Ошибка: ANTHROPIC_API_KEY не найден.\n"
        "Добавьте в .env:\n  ANTHROPIC_API_KEY=sk-ant-..."
    )

MODEL = "claude-opus-4-6"          # Opus для точного парсинга; можно сменить на sonnet
MAX_IMAGES = 3                      # Сколько картинок передавать в Vision (0 = отключить)
IMAGE_MAX_BYTES = 4 * 1024 * 1024  # 4 MB — лимит Anthropic на base64-изображение

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Промпт
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """Ты — эксперт по анализу объявлений коммерческой недвижимости.
Получив текст объявления (и возможно фотографии), извлекаешь строго указанные поля.
Отвечаешь ТОЛЬКО валидным JSON без пояснений, markdown-блоков и лишних символов."""

USER_PROMPT = """Проанализируй текст объявления о недвижимости.
Извлеки следующие данные:

- Price         — цена в числах (только цифры, без валюты; если несколько — бери меньшую)
- Area          — общая площадь в кв.м (только число)
- Object_Type   — тип объекта (строка на украинском или русском языке объявления)
- Renovation    — состояние: одно из ["Без ремонта", "Косметика", "Євро", "Дизайнерський", "Добрий стан", "Потребує ремонту"]
- Year_Built    — год постройки (число) или null если не указан
- Furniture     — есть мебель: true / false
- Location      — краткое описание местоположения (город / район / трасса), строка или null
- Land_Area_ha  — площадь земельного участка в гектарах (число) или null
- Floors        — количество этажей (число) или null

Верни только чистый JSON, например:
{
  "Price": 2500000,
  "Area": 10323,
  "Object_Type": "Складський комплекс",
  "Renovation": "Добрий стан",
  "Year_Built": 1985,
  "Furniture": false,
  "Location": "Київська обл., Бучанський р-н",
  "Land_Area_ha": 2.1,
  "Floors": 1
}

Текст объявления:
"""

# ---------------------------------------------------------------------------
# Парсинг YAML frontmatter
# ---------------------------------------------------------------------------

FRONTMATTER_RE = re.compile(r"^---\n(.*?)\n---\n?(.*)", re.DOTALL)


def parse_md(path: Path) -> tuple[dict, str]:
    """
    Возвращает (frontmatter_dict, body_text).
    Если frontmatter отсутствует — возвращает ({}, полный текст).
    """
    raw = path.read_text(encoding="utf-8")
    m = FRONTMATTER_RE.match(raw)
    if not m:
        return {}, raw
    try:
        fm = yaml.safe_load(m.group(1)) or {}
    except yaml.YAMLError as e:
        log.warning(f"YAML-ошибка в {path.name}: {e}")
        fm = {}
    return fm, m.group(2)


def write_md(path: Path, frontmatter: dict, body: str) -> None:
    """Записывает файл обратно: YAML frontmatter + тело."""
    # allow_unicode=True чтобы кириллица не эскейпилась
    fm_str = yaml.dump(
        frontmatter,
        allow_unicode=True,
        default_flow_style=False,
        sort_keys=False,
    ).rstrip("\n")
    path.write_text(f"---\n{fm_str}\n---\n{body}", encoding="utf-8")


# ---------------------------------------------------------------------------
# Работа с изображениями
# ---------------------------------------------------------------------------

def _find_image_urls(body: str) -> list[str]:
    """Извлекает URL изображений из Markdown-тела заметки."""
    return re.findall(r"!\[.*?\]\((https?://[^\)]+)\)", body)


def _fetch_image_b64(url: str) -> tuple[str, str] | None:
    """
    Скачивает изображение по URL, возвращает (base64_data, media_type).
    Возвращает None при ошибке или если файл слишком большой.
    """
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            content_type = resp.headers.get("Content-Type", "image/jpeg").split(";")[0].strip()
            data = resp.read(IMAGE_MAX_BYTES + 1)
            if len(data) > IMAGE_MAX_BYTES:
                log.debug(f"Изображение слишком большое, пропускаю: {url[:60]}")
                return None
            return base64.standard_b64encode(data).decode(), content_type
    except Exception as e:
        log.debug(f"Не удалось загрузить изображение {url[:60]}: {e}")
        return None


def _build_image_blocks(body: str, max_images: int) -> list[dict]:
    """Формирует список content-блоков с изображениями для Anthropic API."""
    if max_images == 0:
        return []
    blocks = []
    for url in _find_image_urls(body)[:max_images]:
        result = _fetch_image_b64(url)
        if result:
            b64, mime = result
            blocks.append({
                "type": "image",
                "source": {"type": "base64", "media_type": mime, "data": b64},
            })
            log.debug(f"Добавлено изображение ({mime}): {url[:60]}")
    return blocks


# ---------------------------------------------------------------------------
# Вызов Claude API
# ---------------------------------------------------------------------------

def call_claude(text: str, image_blocks: list[dict]) -> dict | None:
    """
    Отправляет запрос к Claude API, возвращает распарсенный JSON-словарь.
    Возвращает None при ошибке.
    """
    import anthropic  # импортируем здесь — после проверки ключа

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Формируем content: сначала изображения (если есть), потом текст промпта
    content: list[dict] = image_blocks + [
        {"type": "text", "text": USER_PROMPT + text[:12_000]}
    ]

    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=512,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": content}],
        )
        raw = response.content[0].text.strip()

        # Убираем ```json ... ``` если модель всё же добавила
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        return json.loads(raw)

    except json.JSONDecodeError as e:
        log.error(f"Не удалось распарсить JSON от Claude: {e}\nОтвет: {raw!r}")
        return None
    except Exception as e:
        log.error(f"Ошибка Claude API: {e}")
        return None


# ---------------------------------------------------------------------------
# Основная логика обработки файла
# ---------------------------------------------------------------------------

# Поля, которые извлекает Claude (используем как allowlist при merge)
EXTRACTED_FIELDS = {
    "Price", "Area", "Object_Type", "Renovation",
    "Year_Built", "Furniture", "Location", "Land_Area_ha", "Floors",
}


def process_file(path: Path, dry_run: bool = False) -> bool:
    """
    Обрабатывает один .md файл.
    Возвращает True если файл был успешно обработан.
    """
    log.info(f"Обрабатываю: {path.name}")

    frontmatter, body = parse_md(path)

    # Собираем текст для анализа: описание из frontmatter + тело заметки
    description = frontmatter.get("description", "")
    title = frontmatter.get("title", "")
    full_text = f"Заголовок: {title}\n\nОписание: {description}\n\n{body}"

    # Загружаем изображения (Vision API)
    image_blocks = _build_image_blocks(body, MAX_IMAGES)
    if image_blocks:
        log.info(f"  Передаю {len(image_blocks)} изображений в Vision API")

    # Вызываем Claude
    extracted = call_claude(full_text, image_blocks)
    if extracted is None:
        log.error(f"  Пропускаю {path.name} — Claude не вернул данные")
        return False

    log.info(f"  Извлечено: {extracted}")

    if dry_run:
        print(f"\n--- {path.name} ---")
        print(json.dumps(extracted, ensure_ascii=False, indent=2))
        return True

    # Merge: Claude перезаписывает только свои поля (не трогает Price/Area если уже есть вручную)
    # Стратегия: Claude-данные имеют приоритет (они свежее ручных)
    for key, value in extracted.items():
        if key in EXTRACTED_FIELDS and value is not None:
            frontmatter[key] = value

    frontmatter["parsed"] = True

    write_md(path, frontmatter, body)
    log.info(f"  ✓ Обновлён: {path.name}")
    return True


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Анализатор клипингов недвижимости через Claude API")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Не записывать файлы — только вывести извлечённый JSON",
    )
    parser.add_argument(
        "--reparse", action="store_true",
        help="Переобрабатывать даже файлы с parsed: true",
    )
    parser.add_argument(
        "--file", type=str, default=None,
        help="Обработать один конкретный файл вместо сканирования Clippings/",
    )
    parser.add_argument(
        "--no-vision", action="store_true",
        help="Отключить Vision API (не загружать изображения)",
    )
    parser.add_argument(
        "--clippings-dir", type=str, default=str(CLIPPINGS_DIR),
        help=f"Путь к папке Clippings (по умолчанию: {CLIPPINGS_DIR})",
    )
    args = parser.parse_args()

    global MAX_IMAGES
    if args.no_vision:
        MAX_IMAGES = 0

    # Определяем список файлов для обработки
    if args.file:
        target = Path(args.file)
        if not target.exists():
            sys.exit(f"Файл не найден: {target}")
        files = [target]
    else:
        clippings_dir = Path(args.clippings_dir)
        if not clippings_dir.is_dir():
            sys.exit(f"Папка не найдена: {clippings_dir}")
        files = sorted(clippings_dir.rglob("*.md"))

    # Фильтрация по parsed: true
    to_process = []
    skipped = 0
    for f in files:
        fm, _ = parse_md(f)
        if fm.get("parsed") is True and not args.reparse:
            skipped += 1
            continue
        to_process.append(f)

    log.info(f"Найдено файлов: {len(files)} | К обработке: {len(to_process)} | Пропущено: {skipped}")

    if not to_process:
        log.info("Нечего обрабатывать.")
        return

    ok = 0
    fail = 0
    for path in to_process:
        success = process_file(path, dry_run=args.dry_run)
        if success:
            ok += 1
        else:
            fail += 1

    log.info(f"Готово. Успешно: {ok} | Ошибок: {fail}")


if __name__ == "__main__":
    main()
