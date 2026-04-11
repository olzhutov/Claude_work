"""
Анализатор клипингов недвижимости — двухступенчатая обработка через Claude.

Методология: .agents/skills/cre-valuation/README.md (разделы B1, C3)

Шаг 1 (classify):  Haiku определяет категорию объекта.
Шаг 2 (extract):   Sonnet извлекает поля по JSON-схеме категории.
Шаг 3 (enrich):    Python нормализует цену, конвертирует UAH→USD,
                    применяет скидку на торг [B1], считает удельные показатели.

Запуск:
    python3 agents/clip_analyzer.py                  # все необработанные в Clippings/
    python3 agents/clip_analyzer.py --dry-run        # без записи, только вывод JSON
    python3 agents/clip_analyzer.py --reparse        # включая parsed: true
    python3 agents/clip_analyzer.py --file path.md   # один файл
    python3 agents/clip_analyzer.py --no-vision      # без загрузки картинок
    python3 agents/clip_analyzer.py --fast           # оба шага на Haiku

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
from typing import Literal, Optional

import yaml
from dotenv import load_dotenv
from pydantic import BaseModel, Field

# ---------------------------------------------------------------------------
# Конфигурация
# ---------------------------------------------------------------------------

BASE_DIR      = Path(__file__).parent.parent
CLIPPINGS_DIR = BASE_DIR / "Clippings"

load_dotenv(BASE_DIR / ".env")

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
if not ANTHROPIC_API_KEY:
    sys.exit("Ошибка: ANTHROPIC_API_KEY не найден в .env")

MODEL_SMART = "claude-sonnet-4-6"          # шаг 2: точное извлечение полей
MODEL_FAST  = "claude-haiku-4-5-20251001"  # шаг 1: классификация + режим --fast

# ── Финансовые константы [B1, C2] ──────────────────────────────────────────

# Жёсткий курс конвертации UAH → USD (менять только здесь)
EXCHANGE_RATE: float = 43.5

# Скидки на торг по типу объекта [B1, Украина 2024-25]
DISCOUNT_RATES: dict[str, float] = {
    "Warehouse":   0.07,   # склады -7%
    "Office":      0.08,   # офисы  -8%
    "Retail":      0.06,   # ритейл -6%
    "Land":        0.05,   # земля  -5%  [ДОПУЩЕНИЕ]
    "Residential": 0.05,   # жильё  -5%  [ДОПУЩЕНИЕ]
}

MAX_IMAGES      = 3
IMAGE_MAX_BYTES = 4 * 1024 * 1024

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# ═══════════════════════════════════════════════════════════════════════════
#  СХЕМЫ КАТЕГОРИЙ (Pydantic)
#  Поля соответствуют структуре карточки объекта [C3]
#  Добавить новую категорию → модель здесь + строка в CATEGORY_REGISTRY
# ═══════════════════════════════════════════════════════════════════════════
# ---------------------------------------------------------------------------


class BaseProperty(BaseModel):
    """Общие поля для всех типов объектов [C3]."""
    Price:          Optional[float] = Field(None, description="Цена числом без валюты и символов")
    Price_Currency: Optional[Literal["USD", "UAH"]] = Field(
        None,
        description=(
            "'USD' если цена в $ / долларах / USD, "
            "'UAH' если в грн / гривнях / ₴. ОБЯЗАТЕЛЬНО."
        ),
    )
    Area:        Optional[float] = Field(None, description="Загальна площа об'єкта, кв.м")
    Location:    Optional[str]   = Field(None, description="Місцезнаходження: місто / район / вулиця")
    Year_Built:  Optional[int]   = Field(None, description="Рік побудови або null")
    Status:      str             = Field("Аналог", description="Аналог | Продаж | Оренда | Закрито")
    Object_Type: Optional[str]   = Field(None, description="Тип об'єкта з оголошення")
    Legal_Status: Optional[str]  = Field(None, description="Право власності / оренда землі / ФДМУ")
    Encumbrances: Optional[str]  = Field(None, description="Обтяження: пам'ятка / застава / арешт або null")


class WarehouseProperty(BaseProperty):
    """Склад / виробництво / логістика / майновий комплекс [C3]."""
    Ceiling_Height: Optional[float] = Field(None, description="Висота стелі, м")
    Power_kW:       Optional[float] = Field(None, description="Електрична потужність, кВт")
    Floor_Type:     Optional[str]   = Field(None, description="Підлога: бетон / асфальт / плитка / насипний")
    Ramps_Docks:    Optional[str]   = Field(None, description="Рампи/доки: є / немає / кількість / тип")
    Land_Area_ha:   Optional[float] = Field(None, description="Площа земельної ділянки, га")
    Floors:         Optional[int]   = Field(None, description="Поверховість")
    Railway:        Optional[bool]  = Field(None, description="Залізнична гілка: true / false")
    Security:       Optional[str]   = Field(None, description="Охорона: цілодобова / відеоспостереження / немає")


class OfficeProperty(BaseProperty):
    """Офіс / бізнес-центр / коворкінг [C3]."""
    Class:          Optional[Literal["A", "B+", "B", "C"]] = Field(None, description="Клас офісу: A / B+ / B / C")
    Layout_Type:    Optional[str]  = Field(None, description="Планування: open-space / кабінетне / змішане")
    Parking_Spaces: Optional[int]  = Field(None, description="Кількість паркомісць")
    Renovation:     Optional[str]  = Field(None, description="Стан: без ремонту / косметичний / євро / дизайнерський")
    Floors:         Optional[int]  = Field(None, description="Поверховість будівлі")
    Management_Co:  Optional[str]  = Field(None, description="Керуюча компанія або null")


class RetailProperty(BaseProperty):
    """Торгівля / стріт-рітейл / ТРЦ / магазин [C3]."""
    Frontage_m:       Optional[float] = Field(None, description="Вітринна лінія / фронтаж, м")
    Floor_in_Building: Optional[int]  = Field(None, description="Поверх розташування")
    Parking_Spaces:   Optional[int]   = Field(None, description="Паркомісць")
    Renovation:       Optional[str]   = Field(None, description="Стан ремонту")
    Separate_Entrance: Optional[bool] = Field(None, description="Окремий вхід: true / false")
    Traffic:          Optional[str]   = Field(None, description="Трафік: пішохідний / автомобільний / змішаний")


class LandProperty(BaseProperty):
    """Земельна ділянка [C3]."""
    Land_Area_ha:        Optional[float] = Field(None, description="Площа, га (пріоритет над Area)")
    Land_Purpose:        Optional[str]   = Field(None, description="Цільове призначення: промисловість / комерція / с/г / житлова")
    Cadastral_Number:    Optional[str]   = Field(None, description="Кадастровий номер якщо вказаний")
    Communications:      Optional[str]   = Field(None, description="Комунікації: електрика / газ / вода / каналізація")
    Distance_to_City_km: Optional[float] = Field(None, description="Відстань до найближчого міста, км")


class ResidentialProperty(BaseProperty):
    """Житлова нерухомість (рідко, але зустрічається в кліпінгах)."""
    Rooms:      Optional[int]  = Field(None, description="Кількість кімнат")
    Floor:      Optional[int]  = Field(None, description="Поверх квартири")
    Renovation: Optional[str]  = Field(None, description="Стан ремонту")
    Furniture:  Optional[bool] = Field(None, description="Меблі: true / false")


# ═══════════════════════════════════════════════════════════════════════════
#  РЕЄСТР КАТЕГОРІЙ
#  key   — ім'я категорії, яке повертає Claude на кроці 1
#  value — Pydantic-модель зі специфічними полями
#  Додати нову категорію: модель вище + рядок тут + правило в CLASSIFY_PROMPT
# ═══════════════════════════════════════════════════════════════════════════

CATEGORY_REGISTRY: dict[str, type[BaseProperty]] = {
    "Warehouse":   WarehouseProperty,
    "Office":      OfficeProperty,
    "Retail":      RetailProperty,
    "Land":        LandProperty,
    "Residential": ResidentialProperty,
}

VALID_CATEGORIES = list(CATEGORY_REGISTRY.keys())
DEFAULT_CATEGORY = "Warehouse"


# ---------------------------------------------------------------------------
# Парсинг / запись YAML frontmatter
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

ALLOWED_MIME = {"image/jpeg", "image/png", "image/gif", "image/webp"}
_EXT_MIME    = {".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                ".png": "image/png",  ".gif": "image/gif", ".webp": "image/webp"}


def _find_image_urls(body: str) -> list[str]:
    return re.findall(r"!\[.*?\]\((https?://[^\)]+)\)", body)


def _fetch_image_b64(url: str) -> tuple[str, str] | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            raw_ct = resp.headers.get("Content-Type", "")
            mime   = raw_ct.split(";")[0].strip().lower()
            if mime not in ALLOWED_MIME:
                from pathlib import PurePosixPath
                ext  = PurePosixPath(url.split("?")[0]).suffix.lower()
                mime = _EXT_MIME.get(ext, "image/jpeg")
            if mime not in ALLOWED_MIME:
                return None
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
            blocks.append({"type": "image",
                           "source": {"type": "base64", "media_type": mime, "data": b64}})
    return blocks


# ---------------------------------------------------------------------------
# Claude API
# ---------------------------------------------------------------------------

def _call_claude(content: list[dict], system: str, model: str,
                 max_tokens: int = 256) -> str | None:
    """Вызов API с автоматическим fallback при ошибке изображений."""
    import anthropic
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    def _req(c: list[dict]) -> str | None:
        try:
            resp = client.messages.create(
                model=model, max_tokens=max_tokens, system=system,
                messages=[{"role": "user", "content": c}],
            )
            return resp.content[0].text.strip()
        except anthropic.BadRequestError as e:
            if "Could not process image" in str(e) or "invalid_request_error" in str(e):
                return "IMAGE_ERROR"
            log.error(f"BadRequest: {e}")
            return None
        except Exception as e:
            log.error(f"Claude API error: {e}")
            return None

    result = _req(content)
    if result == "IMAGE_ERROR":
        text_only = [b for b in content if b.get("type") == "text"]
        if len(text_only) < len(content):
            log.warning("Изображения отклонены API — повтор только с текстом")
            result = _req(text_only)
        if result == "IMAGE_ERROR":
            return None
    return result


def _clean_json(raw: str) -> str:
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return raw.strip()


# ── Шаг 1: Классификация ────────────────────────────────────────────────────

CLASSIFY_SYSTEM = (
    "Ти — експерт з комерційної нерухомості. "
    "Відповідай ТІЛЬКИ одним словом — назвою категорії."
)

CLASSIFY_PROMPT = f"""Визнач категорію об'єкта нерухомості з тексту оголошення.
Вибери ОДНУ категорію зі списку: {", ".join(VALID_CATEGORIES)}.

Правила:
- Warehouse  → склад, ангар, виробництво, логістика, майновий комплекс
- Office     → офіс, бізнес-центр, коворкінг, адміністративне приміщення
- Retail     → магазин, торговий центр, стріт-рітейл, торгове приміщення
- Land       → земельна ділянка, земля, ділянка
- Residential → квартира, будинок, котедж, апартаменти

Текст оголошення:
"""


def classify_category(text: str, model: str) -> str:
    content = [{"type": "text", "text": CLASSIFY_PROMPT + text[:6_000]}]
    raw = _call_claude(content, CLASSIFY_SYSTEM, model=model, max_tokens=20)
    if not raw:
        return DEFAULT_CATEGORY
    for cat in VALID_CATEGORIES:
        if cat.lower() in raw.lower():
            return cat
    log.warning(f"Категория не распознана: {raw!r} → {DEFAULT_CATEGORY}")
    return DEFAULT_CATEGORY


# ── Шаг 2: Извлечение полей ─────────────────────────────────────────────────

EXTRACT_SYSTEM = (
    "Ти — аналітик комерційної нерухомості. "
    "Відповідай ТІЛЬКИ валідним JSON без пояснень і markdown-блоків."
)


def extract_fields(text: str, category: str,
                   image_blocks: list[dict], model: str) -> dict | None:
    model_cls  = CATEGORY_REGISTRY[category]
    schema_json = json.dumps(model_cls.model_json_schema(), ensure_ascii=False, indent=2)

    prompt = (
        f"Категорія об'єкта: {category}\n\n"
        f"Витягни дані з оголошення. JSON Schema:\n\n{schema_json}\n\n"
        "Правила:\n"
        "- Числа без одиниць виміру (тільки цифри)\n"
        "- Поле не вказано в оголошенні → null\n"
        "- Status за замовчуванням: 'Аналог'\n"
        "- Boolean: true або false\n"
        "- Price — тільки число (напр.: 1350000)\n"
        "- Price_Currency — ОБОВ'ЯЗКОВО: 'USD' якщо ціна в $ / доларах, "
        "'UAH' якщо в грн / гривнях / ₴\n\n"
        f"Текст оголошення:\n{text[:12_000]}"
    )

    content: list[dict] = image_blocks + [{"type": "text", "text": prompt}]
    raw = _call_claude(content, EXTRACT_SYSTEM, model=model, max_tokens=1024)
    if not raw:
        return None
    try:
        data = json.loads(_clean_json(raw))
    except json.JSONDecodeError as e:
        log.error(f"JSON parse error: {e}\n→ {raw!r}")
        return None

    try:
        validated = model_cls.model_validate(data)
        return validated.model_dump(exclude_none=True)
    except Exception as e:
        log.warning(f"Pydantic validation: {e} — используется сырой JSON")
        return data


# ---------------------------------------------------------------------------
# Шаг 3: Нормализация цены + расчёт удельных показателей [B1, C2]
# ---------------------------------------------------------------------------

def enrich_price(extracted: dict, category: str) -> dict:
    """
    Конвертирует Price в USD, применяет скидку на торг [B1],
    рассчитывает удельные показатели.

    Добавляет поля:
      Price             — итоговая цена в USD (int)
      Price_Currency    — исходная валюта
      Price_Adjusted    — цена после скидки на торг (int)
      Price_per_sqm     — USD/м² до торга (int)
      Price_per_sqm_Adjusted — USD/м² после торга (int)
    """
    result   = dict(extracted)
    raw_price = result.get("Price")
    currency  = result.get("Price_Currency") or "USD"
    area      = result.get("Area")
    discount  = DISCOUNT_RATES.get(category, 0.0)

    if raw_price is None:
        return result

    # ── Конвертация UAH → USD ───────────────────────────────────────────────
    if currency == "UAH":
        price_usd = raw_price / EXCHANGE_RATE
        log.info(f"  [C2] Конвертация: {raw_price:,.0f} грн ÷ {EXCHANGE_RATE} = {price_usd:,.0f} $")
    else:
        price_usd = float(raw_price)

    result["Price"] = int(round(price_usd))

    # ── Скидка на торг [B1] ─────────────────────────────────────────────────
    price_adj = price_usd * (1 - discount)
    result["Price_Adjusted"] = int(round(price_adj))
    log.info(
        f"  [B1] Торг -{discount*100:.0f}%: "
        f"{result['Price']:,} $ → {result['Price_Adjusted']:,} $"
    )

    # ── Удельные показатели ─────────────────────────────────────────────────
    if area and area > 0:
        result["Price_per_sqm"]          = int(round(price_usd / area))
        result["Price_per_sqm_Adjusted"] = int(round(price_adj / area))
        log.info(
            f"  [B1] Питома: {result['Price_per_sqm']:,} $/м²  "
            f"(після торгу: {result['Price_per_sqm_Adjusted']:,} $/м²)"
        )
    else:
        log.warning("  Area не определена — удельные показатели не рассчитаны")

    return result


# ---------------------------------------------------------------------------
# Основная логика обработки файла
# ---------------------------------------------------------------------------

# Все поля разрешённые к записи в frontmatter
_ALLOWED_FIELDS: set[str] = set()
for _cls in CATEGORY_REGISTRY.values():
    _ALLOWED_FIELDS.update(_cls.model_fields.keys())
_ALLOWED_FIELDS.update({"Price_Adjusted", "Price_per_sqm", "Price_per_sqm_Adjusted"})


def process_file(path: Path, dry_run: bool = False,
                 fast_mode: bool = False) -> dict | None:
    """
    Обрабатывает один .md файл.
    Возвращает словарь с результатами для сводной таблицы или None при ошибке.
    """
    log.info(f"Обрабатываю: {path.name}")

    frontmatter, body = parse_md(path)
    title       = frontmatter.get("title", "")
    description = frontmatter.get("description", "")
    full_text   = f"Заголовок: {title}\n\nОпис: {description}\n\n{body}"

    image_blocks = _build_image_blocks(body, MAX_IMAGES)
    if image_blocks:
        log.info(f"  Vision: {len(image_blocks)} зображень")

    classify_model = MODEL_FAST
    extract_model  = MODEL_FAST if fast_mode else MODEL_SMART

    # Шаг 1
    category = classify_category(full_text, model=classify_model)
    log.info(f"  Категорія → {category}")

    # Шаг 2
    extracted = extract_fields(full_text, category, image_blocks, model=extract_model)
    if extracted is None:
        log.error(f"  Пропускаю {path.name}")
        return None

    log.info(f"  Вилучено (сире): {extracted}")

    # Шаг 3
    enriched = enrich_price(extracted, category)

    # Сводка для итоговой таблицы
    summary = {
        "file":      path.name[:55],
        "category":  category,
        "currency":  enriched.get("Price_Currency", "?"),
        "price_orig": extracted.get("Price"),
        "price_usd":  enriched.get("Price"),
        "price_adj":  enriched.get("Price_Adjusted"),
        "ppsm":       enriched.get("Price_per_sqm"),
        "ppsm_adj":   enriched.get("Price_per_sqm_Adjusted"),
        "area":       enriched.get("Area"),
    }

    if dry_run:
        return summary

    # Запись в frontmatter
    for key, value in enriched.items():
        if key in _ALLOWED_FIELDS and key != "Price_USD_raw":
            frontmatter[key] = value

    frontmatter["Category"] = category
    frontmatter["parsed"]   = True

    write_md(path, frontmatter, body)
    log.info(f"  ✓ Записано: {path.name}")
    return summary


# ---------------------------------------------------------------------------
# Сводная таблица
# ---------------------------------------------------------------------------

def print_summary(results: list[dict]) -> None:
    if not results:
        return

    sep  = "─" * 110
    head = (
        f"{'Файл':<52} {'Кат':>9} {'Вал':>4} "
        f"{'Ціна ориг':>12} {'Ціна USD':>12} {'Торг USD':>12} "
        f"{'$/м²':>7} {'$/м² торг':>10}"
    )

    print(f"\n{'═'*110}")
    print("  ЗВЕДЕНА ТАБЛИЦЯ РЕЗУЛЬТАТІВ ОБРОБКИ")
    print(f"{'═'*110}")
    print(head)
    print(sep)

    for r in results:
        def _fmt(v):
            if v is None:
                return "—"
            return f"{int(v):,}".replace(",", " ")

        orig_str = _fmt(r["price_orig"])
        if r["currency"] == "UAH":
            orig_str += " ₴"
        else:
            orig_str += " $"

        print(
            f"  {r['file']:<50} {r['category']:>9} {r['currency']:>4} "
            f"{orig_str:>13} {_fmt(r['price_usd']):>11}$ "
            f"{_fmt(r['price_adj']):>11}$ "
            f"{_fmt(r['ppsm']):>7} {_fmt(r['ppsm_adj']):>10}"
        )

    print(sep)
    print(f"  Оброблено: {len(results)} файлів\n")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Аналізатор кліпінгів нерухомості [методологія: .agents/skills/cre-valuation/]"
    )
    parser.add_argument("--dry-run",       action="store_true")
    parser.add_argument("--reparse",       action="store_true")
    parser.add_argument("--file",          type=str, default=None)
    parser.add_argument("--no-vision",     action="store_true")
    parser.add_argument("--fast",          action="store_true")
    parser.add_argument("--clippings-dir", type=str, default=str(CLIPPINGS_DIR))
    args = parser.parse_args()

    global MAX_IMAGES
    if args.no_vision:
        MAX_IMAGES = 0

    if args.file:
        target = Path(args.file)
        if not target.exists():
            sys.exit(f"Файл не знайдено: {target}")
        files = [target]
    else:
        d = Path(args.clippings_dir)
        if not d.is_dir():
            sys.exit(f"Папку не знайдено: {d}")
        files = sorted(d.rglob("*.md"))

    to_process, skipped = [], 0
    for f in files:
        fm, _ = parse_md(f)
        if fm.get("parsed") is True and not args.reparse:
            skipped += 1
            continue
        to_process.append(f)

    log.info(f"Файлів: {len(files)} | До обробки: {len(to_process)} | Пропущено: {skipped}")

    if not to_process:
        log.info("Нічого обробляти.")
        return

    results = []
    ok = fail = 0
    for path in to_process:
        r = process_file(path, dry_run=args.dry_run, fast_mode=args.fast)
        if r:
            results.append(r)
            ok += 1
        else:
            fail += 1

    log.info(f"Готово. Успішно: {ok} | Помилок: {fail}")
    print_summary(results)


# ═══════════════════════════════════════════════════════════════════════════
#  ЯК ДОДАТИ НОВУ КАТЕГОРІЮ
# ═══════════════════════════════════════════════════════════════════════════
#
#  1. Створи Pydantic-модель (успадковує BaseProperty):
#
#     class HotelProperty(BaseProperty):
#         Stars:       Optional[int]  = Field(None, description="Зірковість")
#         Rooms_Count: Optional[int]  = Field(None, description="Кількість номерів")
#
#  2. Додай рядок у CATEGORY_REGISTRY:
#         "Hotel": HotelProperty,
#
#  3. Додай скидку в DISCOUNT_RATES:
#         "Hotel": 0.07,
#
#  4. Додай правило в CLASSIFY_PROMPT:
#         - Hotel → готель, хостел, апарт-готель
#
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    main()
