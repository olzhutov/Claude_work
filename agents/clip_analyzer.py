"""
Анализатор клипингов недвижимости — двухступенчатая обработка через Claude.

Шаг 1 (classify):  Claude определяет категорию объекта из текста.
Шаг 2 (extract):   Claude извлекает поля по JSON-схеме, специфичной для категории.
                    Схема генерируется автоматически из Pydantic-модели.

Добавить новую категорию — см. раздел «КАК ДОБАВИТЬ НОВУЮ КАТЕГОРИЮ» в конце файла.

Запуск:
    python3 agents/clip_analyzer.py                      # все необработанные в Clippings/
    python3 agents/clip_analyzer.py --dry-run            # без записи, только вывод JSON
    python3 agents/clip_analyzer.py --reparse            # включая parsed: true
    python3 agents/clip_analyzer.py --file "path.md"     # один файл
    python3 agents/clip_analyzer.py --no-vision          # без загрузки картинок
    python3 agents/clip_analyzer.py --fast               # модель claude-haiku (дешевле)

Зависимости: anthropic, pydantic, PyYAML, python-dotenv
"""

from __future__ import annotations

import argparse
import base64
import json
import logging
import os
import re
import sys
import urllib.request
from pathlib import Path
from typing import Annotated, Literal, Optional

import yaml
from dotenv import load_dotenv
from pydantic import BaseModel, Field

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

MODEL_SMART = "claude-sonnet-4-6"   # для точного извлечения (шаг 2)
MODEL_FAST  = "claude-haiku-4-5-20251001"  # для классификации (шаг 1) и режима --fast

MAX_IMAGES     = 3          # сколько картинок передавать в Vision (0 = отключить)
IMAGE_MAX_BYTES = 4 * 1024 * 1024

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# ═══════════════════════════════════════════════════════════════════════════
#  СХЕМЫ КАТЕГОРИЙ (Pydantic)
#  Добавить новую категорию → только здесь + одна строка в CATEGORY_REGISTRY
# ═══════════════════════════════════════════════════════════════════════════
# ---------------------------------------------------------------------------


class BaseProperty(BaseModel):
    """Общие поля для всех типов объектов."""
    Price:      Optional[float] = Field(None, description="Цена в числах (только цифры, без валюты)")
    Area:       Optional[float] = Field(None, description="Общая площадь объекта, кв.м")
    Location:   Optional[str]   = Field(None, description="Краткое описание местоположения: город / район / трасса")
    Year_Built: Optional[int]   = Field(None, description="Год постройки/сдачи или null")
    Status:     str             = Field("Аналог", description="Статус объекта: Аналог | Продаж | Оренда | Закрито")
    Object_Type: Optional[str]  = Field(None, description="Тип объекта из объявления (строка)")


class WarehouseProperty(BaseProperty):
    """Склад / производство / логистика / имущественный комплекс."""
    Ceiling_Height: Optional[float] = Field(None, description="Высота потолков, м")
    Power_kW:       Optional[float] = Field(None, description="Электрическая мощность, кВт")
    Floor_Type:     Optional[str]   = Field(None, description="Тип пола: бетон / асфальт / плитка / насипний / інше")
    Ramps_Docks:    Optional[str]   = Field(None, description="Наличие рамп/доков: есть / нет / 2 рампи / etc.")
    Land_Area_ha:   Optional[float] = Field(None, description="Площадь земельного участка, га")
    Floors:         Optional[int]   = Field(None, description="Этажность")


class OfficeProperty(BaseProperty):
    """Офис / бизнес-центр / коворкинг."""
    Class:          Optional[Literal["A", "B", "C", "B+"]] = Field(None, description="Класс офиса: A / B+ / B / C")
    Layout_Type:    Optional[str]  = Field(None, description="Тип планировки: open-space / кабинетная / смешанная")
    Parking_Spaces: Optional[int]  = Field(None, description="Количество парковочных мест")
    Renovation:     Optional[str]  = Field(None, description="Ремонт: без ремонту / косметичний / євро / дизайнерський / добрий стан")
    Floors:         Optional[int]  = Field(None, description="Этажность здания")


class RetailProperty(BaseProperty):
    """Торговля / стрит-ритейл / ТРЦ / магазин."""
    Frontage_m:      Optional[float] = Field(None, description="Ширина витрины / фронтаж, м")
    Floor_in_Building: Optional[int] = Field(None, description="Этаж расположения помещения")
    Parking_Spaces:  Optional[int]   = Field(None, description="Парковочных мест")
    Renovation:      Optional[str]   = Field(None, description="Состояние ремонта")
    Separate_Entrance: Optional[bool] = Field(None, description="Есть отдельный вход: true / false")


class LandProperty(BaseProperty):
    """Земельный участок."""
    Land_Area_ha:    Optional[float] = Field(None, description="Площадь участка, га (приоритет над Area)")
    Land_Purpose:    Optional[str]   = Field(None, description="Целевое назначение: промисловість / комерція / с/г / житлова / інше")
    Cadastral_Number: Optional[str]  = Field(None, description="Кадастровый номер, если указан")
    Communications:  Optional[str]   = Field(None, description="Коммуникации: электричество / газ / вода / каналізація")
    Distance_to_City_km: Optional[float] = Field(None, description="Расстояние до ближайшего города, км")


class ResidentialProperty(BaseProperty):
    """Жилая недвижимость (квартиры, дома) — нетипично, но встречается в клипингах."""
    Rooms:       Optional[int]  = Field(None, description="Количество комнат")
    Floor:       Optional[int]  = Field(None, description="Этаж квартиры")
    Renovation:  Optional[str]  = Field(None, description="Состояние ремонта")
    Furniture:   Optional[bool] = Field(None, description="Есть мебель: true / false")


# ═══════════════════════════════════════════════════════════════════════════
#  РЕЕСТР КАТЕГОРИЙ
#  key   — точное имя категории, которое возвращает Claude на шаге 1
#  value — Pydantic-модель с полями для извлечения
#
#  Чтобы добавить новую категорию: создай модель выше и добавь строку здесь.
# ═══════════════════════════════════════════════════════════════════════════

CATEGORY_REGISTRY: dict[str, type[BaseProperty]] = {
    "Warehouse":    WarehouseProperty,
    "Office":       OfficeProperty,
    "Retail":       RetailProperty,
    "Land":         LandProperty,
    "Residential":  ResidentialProperty,
}

VALID_CATEGORIES = list(CATEGORY_REGISTRY.keys())
DEFAULT_CATEGORY = "Warehouse"   # fallback если Claude не уверен


# ---------------------------------------------------------------------------
# Парсинг YAML frontmatter
# ---------------------------------------------------------------------------

FRONTMATTER_RE = re.compile(r"^---\n(.*?)\n---\n?(.*)", re.DOTALL)


def parse_md(path: Path) -> tuple[dict, str]:
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
    fm_str = yaml.dump(
        frontmatter,
        allow_unicode=True,
        default_flow_style=False,
        sort_keys=False,
    ).rstrip("\n")
    path.write_text(f"---\n{fm_str}\n---\n{body}", encoding="utf-8")


# ---------------------------------------------------------------------------
# Vision: загрузка изображений
# ---------------------------------------------------------------------------

def _find_image_urls(body: str) -> list[str]:
    return re.findall(r"!\[.*?\]\((https?://[^\)]+)\)", body)


def _fetch_image_b64(url: str) -> tuple[str, str] | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            mime = resp.headers.get("Content-Type", "image/jpeg").split(";")[0].strip()
            data = resp.read(IMAGE_MAX_BYTES + 1)
            if len(data) > IMAGE_MAX_BYTES:
                return None
            return base64.standard_b64encode(data).decode(), mime
    except Exception as e:
        log.debug(f"Не удалось загрузить {url[:60]}: {e}")
        return None


def _build_image_blocks(body: str, max_images: int) -> list[dict]:
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
    return blocks


# ---------------------------------------------------------------------------
# Claude API: вызовы
# ---------------------------------------------------------------------------

def _call_claude(
    content: list[dict],
    system: str,
    model: str,
    max_tokens: int = 256,
) -> str | None:
    """Низкоуровневый вызов Anthropic Messages API."""
    import anthropic
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=max_tokens,
            system=system,
            messages=[{"role": "user", "content": content}],
        )
        return resp.content[0].text.strip()
    except Exception as e:
        log.error(f"Claude API error: {e}")
        return None


def _clean_json(raw: str) -> str:
    """Убирает markdown-блоки ``` если модель их добавила."""
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return raw.strip()


# ── Шаг 1: Классификация категории ──────────────────────────────────────────

CLASSIFY_SYSTEM = "Ты — эксперт по коммерческой недвижимости. Отвечай только одним словом — названием категории."

CLASSIFY_PROMPT = f"""Определи категорию объекта недвижимости из текста объявления.
Выбери ОДНУ категорию из списка: {", ".join(VALID_CATEGORIES)}.
Отвечай только одним словом — точным именем категории из списка.

Правила выбора:
- Warehouse  → склад, ангар, производство, логистика, имущественный комплекс
- Office     → офис, бизнес-центр, коворкинг, административное помещение
- Retail     → магазин, торговый центр, стрит-ритейл, торговое помещение
- Land       → земельный участок, земля, ділянка
- Residential → квартира, дом, котедж, апартаменты

Текст объявления:
"""


def classify_category(text: str, model: str) -> str:
    """Шаг 1: определяет категорию объекта. Возвращает имя категории."""
    content = [{"type": "text", "text": CLASSIFY_PROMPT + text[:6_000]}]
    raw = _call_claude(content, CLASSIFY_SYSTEM, model=model, max_tokens=20)
    if not raw:
        return DEFAULT_CATEGORY
    # Ищем точное совпадение среди допустимых категорий
    for cat in VALID_CATEGORIES:
        if cat.lower() in raw.lower():
            return cat
    log.warning(f"Категория не распознана: {raw!r} → используется {DEFAULT_CATEGORY}")
    return DEFAULT_CATEGORY


# ── Шаг 2: Извлечение полей по схеме ────────────────────────────────────────

EXTRACT_SYSTEM = (
    "Ты — эксперт по анализу объявлений коммерческой недвижимости. "
    "Отвечай ТОЛЬКО валидным JSON без пояснений и markdown-блоков."
)


def _schema_description(model_cls: type[BaseProperty]) -> str:
    """Формирует читаемое описание полей для промпта из Pydantic-схемы."""
    schema = model_cls.model_json_schema()
    props = schema.get("properties", {})
    lines = []
    for field_name, field_info in props.items():
        desc = field_info.get("description", "")
        field_type = field_info.get("type", field_info.get("anyOf", ""))
        lines.append(f'  "{field_name}": {desc}')
    return "\n".join(lines)


def extract_fields(
    text: str,
    category: str,
    image_blocks: list[dict],
    model: str,
) -> dict | None:
    """Шаг 2: извлекает поля по схеме категории. Возвращает словарь или None."""
    model_cls = CATEGORY_REGISTRY[category]

    # Пример заполненной схемы для промпта
    schema_json = json.dumps(
        model_cls.model_json_schema(),
        ensure_ascii=False,
        indent=2,
    )

    prompt = (
        f"Категория объекта: {category}\n\n"
        f"Извлеки данные из объявления и верни JSON строго по этой JSON Schema:\n\n"
        f"{schema_json}\n\n"
        "Правила:\n"
        "- Числа без единиц измерения (только цифры)\n"
        "- Если поле не указано в объявлении — верни null\n"
        "- Status по умолчанию: 'Аналог' (если не указано иное)\n"
        "- Для boolean: true или false (строчными)\n\n"
        f"Текст объявления:\n{text[:12_000]}"
    )

    content: list[dict] = image_blocks + [{"type": "text", "text": prompt}]
    raw = _call_claude(content, EXTRACT_SYSTEM, model=model, max_tokens=768)
    if not raw:
        return None
    try:
        data = json.loads(_clean_json(raw))
    except json.JSONDecodeError as e:
        log.error(f"JSON parse error: {e}\nОтвет: {raw!r}")
        return None

    # Валидация через Pydantic (приводит типы и отбрасывает лишние поля)
    try:
        validated = model_cls.model_validate(data)
        return validated.model_dump(exclude_none=True)
    except Exception as e:
        log.warning(f"Pydantic validation warning: {e} — используется сырой JSON")
        return data


# ---------------------------------------------------------------------------
# Основная логика обработки файла
# ---------------------------------------------------------------------------

def process_file(
    path: Path,
    dry_run: bool = False,
    fast_mode: bool = False,
) -> bool:
    log.info(f"Обрабатываю: {path.name}")

    frontmatter, body = parse_md(path)

    # Текст для анализа: заголовок + description из frontmatter + тело
    title       = frontmatter.get("title", "")
    description = frontmatter.get("description", "")
    full_text   = f"Заголовок: {title}\n\nОписание: {description}\n\n{body}"

    # Загружаем изображения
    image_blocks = _build_image_blocks(body, MAX_IMAGES)
    if image_blocks:
        log.info(f"  Vision: {len(image_blocks)} изображений")

    classify_model = MODEL_FAST                          # классификация всегда на Haiku
    extract_model  = MODEL_FAST if fast_mode else MODEL_SMART

    # ── Шаг 1: классификация ──────────────────────────────────────────────
    category = classify_category(full_text, model=classify_model)
    log.info(f"  Категория → {category}")

    # ── Шаг 2: извлечение полей ───────────────────────────────────────────
    extracted = extract_fields(full_text, category, image_blocks, model=extract_model)
    if extracted is None:
        log.error(f"  Пропускаю {path.name} — не удалось извлечь данные")
        return False

    log.info(f"  Извлечено: {extracted}")

    if dry_run:
        print(f"\n{'─'*60}")
        print(f"Файл:      {path.name}")
        print(f"Категория: {category}")
        print(json.dumps(extracted, ensure_ascii=False, indent=2))
        return True

    # ── Обновляем frontmatter ─────────────────────────────────────────────
    # Перечень полей которые разрешено перезаписывать из Claude
    allowed_fields: set[str] = set()
    for cls in CATEGORY_REGISTRY.values():
        allowed_fields.update(cls.model_fields.keys())

    for key, value in extracted.items():
        if key in allowed_fields:
            frontmatter[key] = value

    frontmatter["Category"] = category   # записываем определённую категорию
    frontmatter["parsed"]   = True

    write_md(path, frontmatter, body)
    log.info(f"  ✓ Записан: {path.name}")
    return True


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Двухступенчатый анализатор клипингов недвижимости через Claude"
    )
    parser.add_argument("--dry-run",   action="store_true", help="Не записывать файлы, только вывести JSON")
    parser.add_argument("--reparse",   action="store_true", help="Переобрабатывать даже parsed: true")
    parser.add_argument("--file",      type=str, default=None, help="Обработать один конкретный файл")
    parser.add_argument("--no-vision", action="store_true", help="Не загружать изображения")
    parser.add_argument("--fast",      action="store_true", help="Использовать Haiku для обоих шагов (дешевле)")
    parser.add_argument("--clippings-dir", type=str, default=str(CLIPPINGS_DIR))
    args = parser.parse_args()

    global MAX_IMAGES
    if args.no_vision:
        MAX_IMAGES = 0

    # Список файлов
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

    to_process, skipped = [], 0
    for f in files:
        fm, _ = parse_md(f)
        if fm.get("parsed") is True and not args.reparse:
            skipped += 1
            continue
        to_process.append(f)

    log.info(f"Файлов: {len(files)} | К обработке: {len(to_process)} | Пропущено: {skipped}")

    if not to_process:
        log.info("Нечего обрабатывать.")
        return

    ok = fail = 0
    for path in to_process:
        if process_file(path, dry_run=args.dry_run, fast_mode=args.fast):
            ok += 1
        else:
            fail += 1

    log.info(f"Готово. Успешно: {ok} | Ошибок: {fail}")


# ═══════════════════════════════════════════════════════════════════════════
#  КАК ДОБАВИТЬ НОВУЮ КАТЕГОРИЮ
# ═══════════════════════════════════════════════════════════════════════════
#
#  1. Создай Pydantic-модель, наследующую BaseProperty:
#
#     class HotelProperty(BaseProperty):
#         Stars:       Optional[int]  = Field(None, description="Звёздность отеля")
#         Rooms_Count: Optional[int]  = Field(None, description="Количество номеров")
#         Restaurant:  Optional[bool] = Field(None, description="Есть ресторан: true/false")
#
#  2. Зарегистрируй в CATEGORY_REGISTRY (одна строка):
#
#     CATEGORY_REGISTRY: dict[str, type[BaseProperty]] = {
#         ...
#         "Hotel": HotelProperty,   # ← добавить здесь
#     }
#
#  3. Добавь правило в CLASSIFY_PROMPT (одна строка в список):
#
#     - Hotel → готель, хостел, апарт-готель, санаторій
#
#  Всё остальное (JSON-схема для промпта, валидация, запись в YAML) — автоматически.
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    main()
