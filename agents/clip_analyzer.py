"""
Аналізатор кліпінгів нерухомості — фінальна версія.
Методологія: .agents/skills/cre-valuation/README.md [B1, C2, C3]

Кроки обробки:
  1. Haiku  — класифікація категорії (Warehouse / Office / Retail / Land / Residential)
  2. Sonnet — витяг полів за JSON-схемою категорії [C3]
  3. Python — нормалізація ціни (UAH→USD), визначення типу угоди (Sale/Rent),
               зонування (Center/Middle/Periphery/Suburbs), знижка на торг [B1],
               питомі показники

Запуск:
    python3 agents/clip_analyzer.py              # необроблені файли в Clippings/
    python3 agents/clip_analyzer.py --dry-run    # без запису — тільки JSON + таблиця
    python3 agents/clip_analyzer.py --reparse    # включно з parsed: true
    python3 agents/clip_analyzer.py --file path  # один файл
    python3 agents/clip_analyzer.py --fast       # обидва кроки на Haiku
    python3 agents/clip_analyzer.py --no-vision  # без завантаження зображень

Залежності: anthropic, pydantic, PyYAML, python-dotenv
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
# Конфігурація
# ---------------------------------------------------------------------------

BASE_DIR      = Path(__file__).parent.parent
CLIPPINGS_DIR = BASE_DIR / "Clippings"

load_dotenv(BASE_DIR / ".env")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
if not ANTHROPIC_API_KEY:
    sys.exit("Помилка: ANTHROPIC_API_KEY не знайдено в .env")

MODEL_SMART = "claude-sonnet-4-6"
MODEL_FAST  = "claude-haiku-4-5-20251001"

# ── Фінансові константи [B1, C2] ──────────────────────────────────────────

EXCHANGE_RATE: float = 43.5   # UAH → USD (міняти тільки тут)

# Знижки на торг [B1, Україна 2024-25]
DISCOUNT_RATES: dict[str, float] = {
    "Warehouse":   0.07,
    "Office":      0.08,
    "Retail":      0.06,
    "Land":        0.05,
    "Residential": 0.05,
}

MAX_IMAGES      = 3
IMAGE_MAX_BYTES = 4 * 1024 * 1024

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Географічне зонування Києва [B1 — локація ±5-25%]
# ---------------------------------------------------------------------------

_CENTER = {
    "печерський", "печерск", "поділ", "подол",
    "шевченківський", "шевченківськ", "шевченковский",
    "старокиївський", "липки", "бесарабка",
}
_MIDDLE = {
    "голосіївський", "голосіїв", "голосеевск",
    "солом'янський", "соломенск", "соломянськ",
    "оболонський", "оболонь",
    "лук'янівка", "лукьяновка",
}
_SUBURBS_MARKERS = ("область", "обл.", "бучанськ", "броварськ", "київська")


def get_location_zone(district: str | None, location: str | None) -> str:
    """
    Визначає зону локації на основі назви району та адреси.
    Center / Middle / Periphery / Suburbs
    """
    text = " ".join(filter(None, [district, location])).lower()
    if not text:
        return "Unknown"
    if any(m in text for m in _SUBURBS_MARKERS):
        return "Suburbs"
    if any(m in text for m in _CENTER):
        return "Center"
    if any(m in text for m in _MIDDLE):
        return "Middle"
    return "Periphery"


# ---------------------------------------------------------------------------
# ═══════════════════════════════════════════════════════════════════════════
#  СХЕМИ КАТЕГОРІЙ (Pydantic) [C3]
#  Додати нову категорію → модель тут + рядок у CATEGORY_REGISTRY
# ═══════════════════════════════════════════════════════════════════════════
# ---------------------------------------------------------------------------


class BaseProperty(BaseModel):
    """Загальні поля для всіх типів об'єктів."""

    # ── Ціна та тип угоди ─────────────────────────────────────────────────
    Deal_Type:      Optional[Literal["Sale", "Rent"]] = Field(
        None,
        description="Тип угоди: 'Sale' — продаж, 'Rent' — оренда",
    )
    Price:          Optional[float] = Field(
        None,
        description="Ціна продажу — тільки число, без валюти і символів",
    )
    Rent_per_sqm:   Optional[float] = Field(
        None,
        description="Місячна орендна ставка за м² — тільки число ($/м²/міс або грн/м²/міс)",
    )
    Rent_Monthly_Total: Optional[float] = Field(
        None,
        description="Загальна місячна орендна плата за весь об'єкт — тільки число",
    )
    Price_Currency: Optional[Literal["USD", "UAH"]] = Field(
        None,
        description="'USD' якщо ціна в $ / доларах, 'UAH' якщо в грн / гривнях / ₴. ОБОВ'ЯЗКОВО.",
    )

    # ── Об'єкт ────────────────────────────────────────────────────────────
    Area:        Optional[float] = Field(None, description="Загальна площа, кв.м")
    Object_Type: Optional[str]   = Field(None, description="Тип об'єкта з оголошення")
    Year_Built:  Optional[int]   = Field(None, description="Рік побудови або null")
    Floors:      Optional[int]   = Field(None, description="Поверховість будівлі")
    Status:      str             = Field("Аналог", description="Аналог | Продаж | Оренда | Закрито")

    # ── Локація ───────────────────────────────────────────────────────────
    Location:    Optional[str]   = Field(None, description="Адреса: місто / район / вулиця")
    District:    Optional[str]   = Field(
        None,
        description="Назва адміністративного району Києва (напр. 'Шевченківський', 'Печерський')"
    )

    # ── Стан і ремонт ─────────────────────────────────────────────────────
    Condition_Type: Optional[str] = Field(
        None,
        description=(
            "Стан приміщення: одне з — 'після будівельників' / 'під оздоблення' / "
            "'з ремонтом' / 'з меблями'"
        ),
    )
    Renovation_Style: Optional[str] = Field(
        None,
        description=(
            "Стиль ремонту якщо є: одне з — 'радянський' / '2000-ні' / "
            "'новий офісний' / 'дизайнерський'"
        ),
    )

    # ── Паркінг ───────────────────────────────────────────────────────────
    Parking_Open:         Optional[bool] = Field(None, description="Відкрита парковка: true / false")
    Parking_Underground:  Optional[bool] = Field(None, description="Підземний паркінг: true / false")

    # ── Правовий статус ───────────────────────────────────────────────────
    Legal_Status: Optional[str] = Field(None, description="Право власності / оренда землі / ФДМУ")
    Encumbrances: Optional[str] = Field(None, description="Обтяження: пам'ятка / застава / арешт або null")


class WarehouseProperty(BaseProperty):
    """Склад / виробництво / логістика / майновий комплекс [C3]."""
    Ceiling_Height: Optional[float] = Field(None, description="Висота стелі, м")
    Power_kW:       Optional[float] = Field(None, description="Електрична потужність, кВт")
    Floor_Type:     Optional[str]   = Field(None, description="Тип підлоги: бетон / асфальт / плитка / насипний")
    Ramps_Docks:    Optional[str]   = Field(None, description="Рампи/доки: є / немає / кількість / тип")
    Land_Area_ha:   Optional[float] = Field(None, description="Площа земельної ділянки, га")
    Railway:        Optional[bool]  = Field(None, description="Залізнична гілка: true / false")
    Security:       Optional[str]   = Field(None, description="Охорона: цілодобова / відеоспостереження / немає")


class OfficeProperty(BaseProperty):
    """Офіс / бізнес-центр / коворкінг [C3]."""
    Building_Class:      Optional[Literal["A", "B+", "B", "C"]] = Field(
        None, description="Клас будівлі: A / B+ / B / C"
    )
    Layout_Type:         Optional[str]  = Field(None, description="Планування: open-space / кабінетне / змішане")
    Parking_Spaces:      Optional[int]  = Field(None, description="Кількість паркомісць (число)")
    Generator:           Optional[bool] = Field(None, description="Генератор: true / false")
    Shelter:             Optional[bool] = Field(None, description="Укриття / бомбосховище: true / false")
    Distance_to_Metro:   Optional[int]  = Field(None, description="Відстань до метро, хвилин пішки")
    Management_Co:       Optional[str]  = Field(None, description="Керуюча компанія або null")


class RetailProperty(BaseProperty):
    """Торгівля / стріт-рітейл / ТРЦ [C3]."""
    Frontage_m:        Optional[float] = Field(None, description="Вітринна лінія / фронтаж, м")
    Floor_in_Building: Optional[int]   = Field(None, description="Поверх розташування")
    Parking_Spaces:    Optional[int]   = Field(None, description="Паркомісць")
    Separate_Entrance: Optional[bool]  = Field(None, description="Окремий вхід: true / false")
    Traffic:           Optional[str]   = Field(None, description="Трафік: пішохідний / автомобільний / змішаний")


class LandProperty(BaseProperty):
    """Земельна ділянка [C3]."""
    Land_Area_ha:        Optional[float] = Field(None, description="Площа, га")
    Land_Purpose:        Optional[str]   = Field(None, description="Цільове призначення")
    Cadastral_Number:    Optional[str]   = Field(None, description="Кадастровий номер якщо вказаний")
    Communications:      Optional[str]   = Field(None, description="Комунікації: електрика / газ / вода")
    Distance_to_City_km: Optional[float] = Field(None, description="Відстань до міста, км")


class ResidentialProperty(BaseProperty):
    """Житлова нерухомість."""
    Rooms:     Optional[int]  = Field(None, description="Кількість кімнат")
    Floor:     Optional[int]  = Field(None, description="Поверх квартири")
    Furniture: Optional[bool] = Field(None, description="Меблі: true / false")


# ═══════════════════════════════════════════════════════════════════════════
#  РЕЄСТР КАТЕГОРІЙ
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

# ── Маппинг категорій → папки для сортування ──────────────────────────────
CATEGORY_DIRS: dict[str, str] = {
    "Warehouse":   "Warehouses",
    "Office":      "Offices",
    "Retail":      "Retail",
    "Land":        "Land",
    "Residential": "Residential",
}

# Всі поля дозволені до запису у frontmatter
_ALLOWED_FIELDS: set[str] = set()
for _cls in CATEGORY_REGISTRY.values():
    _ALLOWED_FIELDS.update(_cls.model_fields.keys())
_ALLOWED_FIELDS.update({
    "Price_Adjusted", "Price_per_sqm", "Price_per_sqm_Adjusted",
    "Rent_Adjusted", "Location_Zone",
})

# ---------------------------------------------------------------------------
# Парсинг / запис YAML frontmatter
# ---------------------------------------------------------------------------

FRONTMATTER_RE = re.compile(r"^---\n(.*?)\n---\n?(.*)", re.DOTALL)


def parse_md(path: Path) -> tuple[dict, str]:
    raw = path.read_text(encoding="utf-8")
    m   = FRONTMATTER_RE.match(raw)
    if not m:
        return {}, raw
    try:
        fm = yaml.safe_load(m.group(1)) or {}
    except yaml.YAMLError as e:
        log.warning(f"YAML-помилка в {path.name}: {e}")
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
# Vision
# ---------------------------------------------------------------------------

ALLOWED_MIME = {"image/jpeg", "image/png", "image/gif", "image/webp"}
_EXT_MIME = {".jpg": "image/jpeg", ".jpeg": "image/jpeg",
             ".png": "image/png", ".gif": "image/gif", ".webp": "image/webp"}


def _find_image_urls(body: str) -> list[str]:
    return re.findall(r"!\[.*?\]\((https?://[^\)]+)\)", body)


def _fetch_image_b64(url: str) -> tuple[str, str] | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            raw_ct = resp.headers.get("Content-Type", "")
            mime   = raw_ct.split(";")[0].strip().lower()
            if mime not in ALLOWED_MIME:
                ext  = Path(url.split("?")[0]).suffix.lower()
                mime = _EXT_MIME.get(ext, "image/jpeg")
            if mime not in ALLOWED_MIME:
                return None
            data = resp.read(IMAGE_MAX_BYTES + 1)
            if len(data) > IMAGE_MAX_BYTES:
                return None
            return base64.standard_b64encode(data).decode(), mime
    except Exception as e:
        log.debug(f"Не вдалось завантажити {url[:60]}: {e}")
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
            log.warning("Зображення відхилені API — повтор тільки з текстом")
            result = _req(text_only)
        if result == "IMAGE_ERROR":
            return None
    return result


def _clean_json(raw: str) -> str:
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return raw.strip()


# ── Крок 1: Класифікація ────────────────────────────────────────────────────

CLASSIFY_SYSTEM = (
    "Ти — експерт з комерційної нерухомості. "
    "Відповідай ТІЛЬКИ одним словом — назвою категорії."
)

CLASSIFY_PROMPT = f"""Визнач категорію об'єкта нерухомості.
Вибери ОДНУ категорію: {", ".join(VALID_CATEGORIES)}.

- Warehouse  → склад, ангар, виробництво, логістика, майновий комплекс
- Office     → офіс, бізнес-центр, коворкінг, адміністративне приміщення
- Retail     → магазин, торговий центр, стріт-рітейл, торгове приміщення
- Land       → земельна ділянка, земля
- Residential → квартира, будинок, котедж

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
    log.warning(f"Категорія не розпізнана: {raw!r} → {DEFAULT_CATEGORY}")
    return DEFAULT_CATEGORY


# ── Крок 2: Витяг полів ─────────────────────────────────────────────────────

EXTRACT_SYSTEM = (
    "Ти — Senior CRE аналітик (Україна). "
    "Відповідай ТІЛЬКИ валідним JSON без пояснень і markdown-блоків."
)


def extract_fields(text: str, category: str,
                   image_blocks: list[dict], model: str) -> dict | None:
    model_cls   = CATEGORY_REGISTRY[category]
    schema_json = json.dumps(model_cls.model_json_schema(), ensure_ascii=False, indent=2)

    prompt = (
        f"Категорія: {category}\n\n"
        f"Витягни дані з оголошення. JSON Schema:\n\n{schema_json}\n\n"
        "Правила:\n"
        "- Числа без одиниць виміру\n"
        "- Відсутнє поле → null\n"
        "- Deal_Type: 'Sale' якщо продаж, 'Rent' якщо оренда\n"
        "- Price: ціна продажу числом\n"
        "- Rent_per_sqm: місячна ставка за м², тільки число\n"
        "- Rent_Monthly_Total: загальна місячна плата, тільки число\n"
        "- Price_Currency: 'USD' або 'UAH' — ОБОВ'ЯЗКОВО\n"
        "- District: назва адміністративного району Києва або null\n"
        "- Parking_Open / Parking_Underground: true/false\n"
        "- Generator / Shelter (Office): true/false\n\n"
        f"Текст:\n{text[:12_000]}"
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
        log.warning(f"Pydantic validation: {e}")
        return data


# ---------------------------------------------------------------------------
# Крок 3: Збагачення — нормалізація ціни, зонування, знижки [B1, C2]
# ---------------------------------------------------------------------------

def enrich(extracted: dict, category: str) -> dict:
    """
    • UAH → USD [C2]
    • Location_Zone (Center / Middle / Periphery / Suburbs)
    • Знижка на торг [B1]: Price_Adjusted або Rent_Adjusted
    • Питомі показники: Price_per_sqm, Price_per_sqm_Adjusted
    """
    r        = dict(extracted)
    currency = r.get("Price_Currency") or "USD"
    area     = r.get("Area")
    discount = DISCOUNT_RATES.get(category, 0.05)
    deal     = r.get("Deal_Type", "Sale")

    def to_usd(val: float | None) -> float | None:
        if val is None:
            return None
        if currency == "UAH":
            usd = val / EXCHANGE_RATE
            log.info(f"  [C2] {val:,.0f} грн ÷ {EXCHANGE_RATE} = {usd:,.0f} $")
            return usd
        return float(val)

    # ── Зонування ──────────────────────────────────────────────────────────
    zone = get_location_zone(r.get("District"), r.get("Location"))
    r["Location_Zone"] = zone

    # ── Sale ────────────────────────────────────────────────────────────────
    if deal == "Sale" and r.get("Price") is not None:
        price_usd = to_usd(r["Price"])
        r["Price"] = int(round(price_usd))

        adj = price_usd * (1 - discount)
        r["Price_Adjusted"] = int(round(adj))
        log.info(f"  [B1] Торг -{discount*100:.0f}%: {r['Price']:,}$ → {r['Price_Adjusted']:,}$")

        if area and area > 0:
            r["Price_per_sqm"]          = int(round(price_usd / area))
            r["Price_per_sqm_Adjusted"] = int(round(adj / area))
            log.info(
                f"  [B1] $/м²: {r['Price_per_sqm']:,} → {r['Price_per_sqm_Adjusted']:,} (після торгу)"
            )

    # ── Rent ─────────────────────────────────────────────────────────────────
    if deal == "Rent":
        # Конвертуємо загальну ставку
        if r.get("Rent_Monthly_Total") is not None:
            r["Rent_Monthly_Total"] = int(round(to_usd(r["Rent_Monthly_Total"])))

        # Якщо Rent_per_sqm відсутній, але є Rent_Monthly_Total + Area — рахуємо
        if r.get("Rent_per_sqm") is None and r.get("Rent_Monthly_Total") and area and area > 0:
            r["Rent_per_sqm"] = round(r["Rent_Monthly_Total"] / area, 2)
            log.info(
                f"  [C3] Rent_per_sqm = {r['Rent_Monthly_Total']:,}$ ÷ {area:.0f}м² "
                f"= {r['Rent_per_sqm']} $/м²/міс"
            )

        if r.get("Rent_per_sqm") is not None:
            rent_usd = to_usd(r["Rent_per_sqm"])
            r["Rent_per_sqm"]  = round(rent_usd, 2)
            r["Rent_Adjusted"] = round(rent_usd * (1 - discount), 2)
            log.info(
                f"  [B1] Оренда торг -{discount*100:.0f}%: "
                f"{r['Rent_per_sqm']} $/м²/міс → {r['Rent_Adjusted']} $/м²/міс"
            )

    log.info(f"  Зона: {zone}")
    return r


# ---------------------------------------------------------------------------
# Обробка файлу
# ---------------------------------------------------------------------------

def process_file(path: Path, dry_run: bool = False,
                 fast_mode: bool = False) -> dict | None:
    log.info(f"Обробляю: {path.name}")

    frontmatter, body = parse_md(path)
    title       = frontmatter.get("title", "")
    description = frontmatter.get("description", "")
    full_text   = f"Заголовок: {title}\n\nОпис: {description}\n\n{body}"

    image_blocks = _build_image_blocks(body, MAX_IMAGES)
    if image_blocks:
        log.info(f"  Vision: {len(image_blocks)} зображень")

    # Крок 1
    category = classify_category(full_text, model=MODEL_FAST)
    log.info(f"  Категорія → {category}")

    # Крок 2
    extract_model = MODEL_FAST if fast_mode else MODEL_SMART
    extracted = extract_fields(full_text, category, image_blocks, extract_model)
    if extracted is None:
        log.error(f"  Пропускаю: {path.name}")
        return None

    log.info(f"  Витягнено (сире): {extracted}")

    # Крок 3
    enriched = enrich(extracted, category)

    # Рядок для зведеної таблиці
    summary = {
        "file":       path.name[:55],
        "category":   category,
        "deal":       enriched.get("Deal_Type", "?"),
        "district":   enriched.get("District", "—"),
        "zone":       enriched.get("Location_Zone", "—"),
        "currency":   enriched.get("Price_Currency", "?"),
        # Sale
        "price_orig": extracted.get("Price"),
        "price_usd":  enriched.get("Price"),
        "price_adj":  enriched.get("Price_Adjusted"),
        "ppsm_adj":   enriched.get("Price_per_sqm_Adjusted"),
        # Rent
        "rent_psm":   enriched.get("Rent_per_sqm"),
        "rent_adj":   enriched.get("Rent_Adjusted"),
        "rent_total": enriched.get("Rent_Monthly_Total"),
    }

    if dry_run:
        return summary

    # Запис у frontmatter
    for key, value in enriched.items():
        if key in _ALLOWED_FIELDS:
            frontmatter[key] = value

    frontmatter["Category"] = category
    frontmatter["parsed"]   = True

    write_md(path, frontmatter, body)
    log.info(f"  ✓ Записано: {path.name}")

    # ── Сортування у підпапку за категорією ───────────────────────────────
    subdir_name = CATEGORY_DIRS.get(category)
    if subdir_name:
        dest_dir = path.parent / subdir_name
        dest_dir.mkdir(exist_ok=True)
        dest = dest_dir / path.name
        path.rename(dest)
        log.info(f"  → Переміщено: {subdir_name}/{path.name}")
        summary["moved_to"] = str(dest_dir.name)

    return summary


# ---------------------------------------------------------------------------
# Зведена таблиця
# ---------------------------------------------------------------------------

def _fmt(v, suffix="") -> str:
    if v is None:
        return "—"
    if isinstance(v, float):
        return f"{v:,.1f}{suffix}"
    return f"{int(v):,}{suffix}".replace(",", " ")


def print_summary(results: list[dict]) -> None:
    if not results:
        return

    W = 120
    print(f"\n{'═' * W}")
    print("  ЗВЕДЕНА ТАБЛИЦЯ РЕЗУЛЬТАТІВ")
    print(f"{'═' * W}")
    print(
        f"  {'Файл':<44} {'Кат':>9} {'Угода':>5} {'Район':<16} {'Зона':>10} "
        f"{'Ціна/Ставка USD':>17} {'Після торгу':>13} {'Папка':>12}"
    )
    print(f"  {'─' * (W - 2)}")

    for r in results:
        deal = r.get("deal", "?")
        if deal == "Sale":
            price_str = _fmt(r["price_usd"], " $")
            adj_str   = _fmt(r["price_adj"], " $")
        else:
            price_str = _fmt(r["rent_psm"], " $/м²")
            adj_str   = _fmt(r["rent_adj"], " $/м²")

        moved = r.get("moved_to", "—")

        print(
            f"  {r['file']:<44} {r['category']:>9} {deal:>5} "
            f"{(r['district'] or '—'):<16} {r['zone']:>10} "
            f"{price_str:>17} {adj_str:>13} {moved:>12}"
        )

    print(f"  {'─' * (W - 2)}")
    print(f"  Оброблено: {len(results)} файлів\n")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="CRE кліпінг-аналізатор [.agents/skills/cre-valuation/]"
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

    results, ok, fail = [], 0, 0
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
#  1. Pydantic-модель (успадковує BaseProperty):
#       class HotelProperty(BaseProperty):
#           Stars: Optional[int] = Field(None, description="Зірковість")
#
#  2. CATEGORY_REGISTRY:   "Hotel": HotelProperty
#  3. DISCOUNT_RATES:      "Hotel": 0.07
#  4. CLASSIFY_PROMPT:     - Hotel → готель, хостел, апарт-готель
#
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    main()
