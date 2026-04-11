"""
Аналізатор кліпінгів нерухомості — оптимізована версія (Haiku-only, no Vision).
Методологія: .agents/skills/cre-valuation/README.md [B1, C2, C3]

Кроки обробки:
  1. Haiku  — класифікація категорії (Warehouse / Office / Retail / Land / Residential)
  2. Haiku  — витяг полів за JSON-схемою категорії [C3]
  3. Python — нормалізація ціни (UAH→USD), визначення типу угоди (Sale/Rent),
               зонування (Center/Middle/Periphery/Suburbs), знижка на торг [B1],
               питомі показники

Запуск:
    python3 agents/clip_analyzer.py              # необроблені файли в Clippings/
    python3 agents/clip_analyzer.py --dry-run    # без запису — тільки JSON + таблиця
    python3 agents/clip_analyzer.py --reparse    # включно з parsed: true
    python3 agents/clip_analyzer.py --file path  # один файл

Залежності: anthropic, pydantic, PyYAML, python-dotenv
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
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

MODEL = "claude-haiku-4-5-20251001"
TEXT_LIMIT = 3500  # символів після очищення

# ── Фінансові константи [B1, C2] ──────────────────────────────────────────

EXCHANGE_RATE: float = 43.5   # UAH → USD (міняти тільки тут)

DISCOUNT_RATES: dict[str, float] = {
    "Warehouse":   0.07,
    "Office":      0.08,
    "Retail":      0.06,
    "Land":        0.05,
    "Residential": 0.05,
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Географічне зонування Києва [B1]
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
# Схеми категорій (Pydantic) [C3]
# ---------------------------------------------------------------------------

class BaseProperty(BaseModel):
    Deal_Type:          Optional[Literal["Sale", "Rent"]] = Field(None, description="Sale або Rent")
    Price:              Optional[float] = Field(None, description="Ціна продажу, число")
    Rent_per_sqm:       Optional[float] = Field(None, description="Орендна ставка $/м²/міс, число")
    Rent_Monthly_Total: Optional[float] = Field(None, description="Загальна місячна оренда, число")
    Price_Currency:     Optional[Literal["USD", "UAH"]] = Field(None, description="USD або UAH")
    Area:               Optional[float] = Field(None, description="Площа, м²")
    Object_Type:        Optional[str]   = Field(None, description="Тип об'єкта")
    Year_Built:         Optional[int]   = Field(None, description="Рік побудови")
    Floors:             Optional[int]   = Field(None, description="Поверховість")
    Status:             str             = Field("Аналог")
    Location:           Optional[str]   = Field(None, description="Адреса")
    District:           Optional[str]   = Field(None, description="Район Києва")
    Condition_Type:     Optional[str]   = Field(None, description="після будівельників / під оздоблення / з ремонтом / з меблями")
    Renovation_Style:   Optional[str]   = Field(None, description="радянський / 2000-ні / новий офісний / дизайнерський")
    Parking_Open:       Optional[bool]  = Field(None, description="Відкрита парковка")
    Parking_Underground:Optional[bool]  = Field(None, description="Підземний паркінг")
    Legal_Status:       Optional[str]   = Field(None, description="Право власності")
    Encumbrances:       Optional[str]   = Field(None, description="Обтяження або null")


class WarehouseProperty(BaseProperty):
    Ceiling_Height: Optional[float] = Field(None, description="Висота стелі, м")
    Power_kW:       Optional[float] = Field(None, description="Потужність, кВт")
    Floor_Type:     Optional[str]   = Field(None, description="бетон / асфальт / плитка / насипний")
    Ramps_Docks:    Optional[str]   = Field(None, description="Рампи/доки")
    Land_Area_ha:   Optional[float] = Field(None, description="Площа ділянки, га")
    Railway:        Optional[bool]  = Field(None, description="Залізнична гілка")
    Security:       Optional[str]   = Field(None, description="Охорона")


class OfficeProperty(BaseProperty):
    Building_Class:    Optional[Literal["A", "B+", "B", "C"]] = Field(None, description="Клас A/B+/B/C")
    Layout_Type:       Optional[str]  = Field(None, description="open-space / кабінетне / змішане")
    Parking_Spaces:    Optional[int]  = Field(None, description="Паркомісць")
    Generator:         Optional[bool] = Field(None, description="Генератор")
    Shelter:           Optional[bool] = Field(None, description="Укриття/бомбосховище")
    Distance_to_Metro: Optional[int]  = Field(None, description="До метро, хв пішки")
    Management_Co:     Optional[str]  = Field(None, description="Керуюча компанія")


class RetailProperty(BaseProperty):
    Frontage_m:        Optional[float] = Field(None, description="Фронтаж, м")
    Floor_in_Building: Optional[int]   = Field(None, description="Поверх")
    Parking_Spaces:    Optional[int]   = Field(None, description="Паркомісць")
    Separate_Entrance: Optional[bool]  = Field(None, description="Окремий вхід")
    Traffic:           Optional[str]   = Field(None, description="пішохідний / автомобільний / змішаний")


class LandProperty(BaseProperty):
    Land_Area_ha:        Optional[float] = Field(None, description="Площа, га")
    Land_Purpose:        Optional[str]   = Field(None, description="Цільове призначення")
    Cadastral_Number:    Optional[str]   = Field(None, description="Кадастровий номер")
    Communications:      Optional[str]   = Field(None, description="Комунікації")
    Distance_to_City_km: Optional[float] = Field(None, description="До міста, км")


class ResidentialProperty(BaseProperty):
    Rooms:     Optional[int]  = Field(None, description="Кімнат")
    Floor:     Optional[int]  = Field(None, description="Поверх")
    Furniture: Optional[bool] = Field(None, description="Меблі")


CATEGORY_REGISTRY: dict[str, type[BaseProperty]] = {
    "Warehouse":   WarehouseProperty,
    "Office":      OfficeProperty,
    "Retail":      RetailProperty,
    "Land":        LandProperty,
    "Residential": ResidentialProperty,
}

VALID_CATEGORIES = list(CATEGORY_REGISTRY.keys())
DEFAULT_CATEGORY = "Warehouse"

CATEGORY_DIRS: dict[str, str] = {
    "Warehouse":   "Warehouses",
    "Office":      "Offices",
    "Retail":      "Retail",
    "Land":        "Land",
    "Residential": "Residential",
}

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
# Очищення тексту
# ---------------------------------------------------------------------------

_MD_NOISE = re.compile(
    r"!\[.*?\]\(.*?\)"       # images
    r"|<[^>]+>"              # HTML tags
    r"|```[\s\S]*?```"       # code blocks
    r"|\[([^\]]+)\]\([^\)]+\)"  # links → keep label
    r"|#{1,6}\s+"            # headings markers
    r"|[*_`]{1,3}"           # bold/italic/code markers
    r"|\|[-:]+\|(?:[-:| ]+\|)*"  # table separators
)


def _clean_text(text: str, limit: int = TEXT_LIMIT) -> str:
    """Видаляє Markdown-розмітку, зайві пробіли і обрізає до limit символів."""
    text = _MD_NOISE.sub(lambda m: m.group(1) or "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r" {2,}", " ", text)
    return text[:limit].strip()


# ---------------------------------------------------------------------------
# Claude API з Prompt Caching
# ---------------------------------------------------------------------------

def _call_claude(text_prompt: str, system_text: str, max_tokens: int = 256) -> str | None:
    """
    Викликає Claude Haiku з prompt caching на системному промпті.
    system_text кешується через cache_control: ephemeral.
    """
    import anthropic
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    try:
        resp = client.messages.create(
            model=MODEL,
            max_tokens=max_tokens,
            system=[
                {
                    "type": "text",
                    "text": system_text,
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            messages=[{"role": "user", "content": text_prompt}],
        )
        return resp.content[0].text.strip()
    except Exception as e:
        log.error(f"Claude API error: {e}")
        return None


def _clean_json(raw: str) -> str:
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return raw.strip()


# ── Крок 1: Класифікація ────────────────────────────────────────────────────

CLASSIFY_SYSTEM = (
    "Ти — CRE аналітик. Відповідай ТІЛЬКИ одним словом — назвою категорії: "
    + ", ".join(VALID_CATEGORIES) + "."
)

CLASSIFY_PROMPT = (
    "Warehouse=склад/ангар/виробництво/логістика/майновий комплекс\n"
    "Office=офіс/бізнес-центр/коворкінг\n"
    "Retail=магазин/ТРЦ/стріт-рітейл\n"
    "Land=земельна ділянка\n"
    "Residential=квартира/будинок\n\n"
    "Визнач категорію:\n"
)


# Примусові keyword-правила (до виклику API) [C3]
_KEYWORD_OVERRIDES: list[tuple[list[str], str]] = [
    (["офіс", "оренда приміщення", "бізнес-центр", "коворкінг"], "Office"),
    (["склад", "ангар", "виробнич", "логістик", "майновий комплекс"], "Warehouse"),
    (["магазин", "трц", "стріт-рітейл", "торговель"], "Retail"),
    (["земельн", "ділянка", "кадастр"], "Land"),
    (["квартир", "будинок", "резиденц"], "Residential"),
]


def _keyword_category(text: str) -> str | None:
    """Повертає категорію якщо є чіткий keyword-збіг, інакше None."""
    lower = text.lower()
    for keywords, cat in _KEYWORD_OVERRIDES:
        if any(kw in lower for kw in keywords):
            return cat
    return None


def classify_category(text: str) -> str:
    clean = _clean_text(text, limit=1500)
    # Спочатку — детермінований keyword-матч (без витрат API)
    kw_cat = _keyword_category(clean)
    if kw_cat:
        log.info(f"  Категорія (keyword) → {kw_cat}")
        return kw_cat
    # Якщо keyword не спрацював — запит до Claude
    raw = _call_claude(CLASSIFY_PROMPT + clean, CLASSIFY_SYSTEM, max_tokens=10)
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
    "Відповідай ТІЛЬКИ валідним JSON без пояснень і markdown-блоків. "
    "Відсутнє поле → null. Числа без одиниць."
)


def extract_fields(text: str, category: str) -> dict | None:
    model_cls   = CATEGORY_REGISTRY[category]
    schema_json = json.dumps(model_cls.model_json_schema(), ensure_ascii=False)
    clean       = _clean_text(text)

    prompt = (
        f"Категорія: {category}. Schema: {schema_json}\n\n"
        "Правила: Deal_Type=Sale/Rent; Price_Currency=USD/UAH (обов'язково); "
        "District=адм. район Києва або null; Parking_*/Generator/Shelter/Railway=true/false.\n\n"
        f"Оголошення:\n{clean}"
    )

    raw = _call_claude(prompt, EXTRACT_SYSTEM, max_tokens=800)
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
# Крок 3: Збагачення [B1, C2]
# ---------------------------------------------------------------------------

_DEAL_TYPE_MAP: dict[str, str] = {
    "sale": "Sale", "продаж": "Sale", "продажа": "Sale",
    "rent": "Rent", "оренда": "Rent", "аренда": "Rent",
}


def _normalize_deal_type(raw: str | None) -> str:
    """Нормалізує Deal_Type до суворих значень Sale/Rent."""
    if not raw:
        return "Sale"
    lower = str(raw).lower().strip()
    return _DEAL_TYPE_MAP.get(lower, "Sale")


def enrich(extracted: dict, category: str) -> dict:
    r        = dict(extracted)
    currency = r.get("Price_Currency") or "USD"
    area     = r.get("Area")
    discount = DISCOUNT_RATES.get(category, 0.05)

    # ── Суворе примусове значення Object_Type (тільки англійські категорії) ──
    r["Object_Type"] = category

    # ── Нормалізація Deal_Type: тільки Sale або Rent ─────────────────────────
    raw_deal = r.get("Deal_Type")
    deal = _normalize_deal_type(raw_deal)
    r["Deal_Type"] = deal

    def to_usd(val: float | None) -> float | None:
        if val is None:
            return None
        if currency == "UAH":
            usd = val / EXCHANGE_RATE
            log.info(f"  [C2] {val:,.0f} грн → {usd:,.0f} $")
            return usd
        return float(val)

    zone = get_location_zone(r.get("District"), r.get("Location"))
    r["Location_Zone"] = zone

    if deal == "Sale" and r.get("Price") is not None:
        price_usd = to_usd(r["Price"])
        r["Price"] = int(round(price_usd))
        adj = price_usd * (1 - discount)
        r["Price_Adjusted"] = int(round(adj))
        log.info(f"  [B1] -{discount*100:.0f}%: {r['Price']:,}$ → {r['Price_Adjusted']:,}$")
        if area and area > 0:
            r["Price_per_sqm"]          = int(round(price_usd / area))
            r["Price_per_sqm_Adjusted"] = int(round(adj / area))

    if deal == "Rent":
        if r.get("Rent_Monthly_Total") is not None:
            r["Rent_Monthly_Total"] = int(round(to_usd(r["Rent_Monthly_Total"])))
        if r.get("Rent_per_sqm") is None and r.get("Rent_Monthly_Total") and area and area > 0:
            r["Rent_per_sqm"] = round(r["Rent_Monthly_Total"] / area, 2)
        if r.get("Rent_per_sqm") is not None:
            rent_usd = to_usd(r["Rent_per_sqm"])
            r["Rent_per_sqm"]  = round(rent_usd, 2)
            r["Rent_Adjusted"] = round(rent_usd * (1 - discount), 2)
            log.info(f"  [B1] Оренда -{discount*100:.0f}%: {r['Rent_per_sqm']} → {r['Rent_Adjusted']} $/м²/міс")

    log.info(f"  Зона: {zone}")
    return r


# ---------------------------------------------------------------------------
# Обробка файлу
# ---------------------------------------------------------------------------

def process_file(path: Path, dry_run: bool = False) -> dict | None:
    log.info(f"Обробляю: {path.name}")

    frontmatter, body = parse_md(path)
    title       = frontmatter.get("title", "")
    description = frontmatter.get("description", "")
    full_text   = f"{title}\n{description}\n{body}"

    category = classify_category(full_text)
    log.info(f"  Категорія → {category}")

    extracted = extract_fields(full_text, category)
    if extracted is None:
        log.error(f"  Пропускаю: {path.name}")
        return None

    log.info(f"  Витягнено: {extracted}")
    enriched = enrich(extracted, category)

    summary = {
        "file":       path.name[:55],
        "category":   category,
        "deal":       enriched.get("Deal_Type", "?"),
        "district":   enriched.get("District", "—"),
        "zone":       enriched.get("Location_Zone", "—"),
        "currency":   enriched.get("Price_Currency", "?"),
        "price_orig": extracted.get("Price"),
        "price_usd":  enriched.get("Price"),
        "price_adj":  enriched.get("Price_Adjusted"),
        "ppsm_adj":   enriched.get("Price_per_sqm_Adjusted"),
        "rent_psm":   enriched.get("Rent_per_sqm"),
        "rent_adj":   enriched.get("Rent_Adjusted"),
        "rent_total": enriched.get("Rent_Monthly_Total"),
    }

    if dry_run:
        return summary

    for key, value in enriched.items():
        if key in _ALLOWED_FIELDS:
            frontmatter[key] = value

    frontmatter["Category"] = category
    frontmatter["parsed"]   = True
    write_md(path, frontmatter, body)
    log.info(f"  ✓ Записано: {path.name}")

    subdir_name = CATEGORY_DIRS.get(category)
    if subdir_name:
        # Цільова папка ЗАВЖДИ відносно кореня CLIPPINGS_DIR, не path.parent
        dest_dir = CLIPPINGS_DIR / subdir_name
        dest_dir.mkdir(exist_ok=True)
        dest = dest_dir / path.name

        if path.resolve() == dest.resolve():
            log.info(f"  ↷ Вже в цільовій папці: {subdir_name}/{path.name}")
            summary["moved_to"] = subdir_name
        elif dest.exists():
            log.warning(f"  ⚠ Файл вже існує в {dest} — пропускаю переміщення")
            summary["moved_to"] = subdir_name
        else:
            path.rename(dest)
            log.info(f"  → Переміщено: {subdir_name}/{path.name}")
            summary["moved_to"] = subdir_name

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

# ---------------------------------------------------------------------------
# Очищення вкладених дублюючих папок
# ---------------------------------------------------------------------------

def cleanup_nested_dirs(clippings_dir: Path) -> None:
    """
    Знаходить файли у Clippings/{Cat}/{Cat}/ (подвійна вкладеність),
    переміщує їх на рівень вище до Clippings/{Cat}/ і видаляє порожні папки.
    Викликається автоматично на початку кожного запуску.
    """
    moved_total = 0

    for cat_dir in sorted(clippings_dir.iterdir()):
        if not cat_dir.is_dir():
            continue
        # Шукаємо підпапки з тією ж назвою або з будь-якою категорійною назвою
        for sub_dir in sorted(cat_dir.iterdir()):
            if not sub_dir.is_dir():
                continue
            # Якщо підпапка є будь-якою з категорійних — це дубль
            all_cat_dirs = set(CATEGORY_DIRS.values())
            if sub_dir.name not in all_cat_dirs:
                continue

            # Переміщуємо *.md з sub_dir → cat_dir або далі в CLIPPINGS_DIR/{sub_dir.name}
            correct_dir = clippings_dir / sub_dir.name
            correct_dir.mkdir(exist_ok=True)

            for md_file in sorted(sub_dir.glob("*.md")):
                dest = correct_dir / md_file.name
                if dest.exists():
                    log.warning(f"  [cleanup] Файл вже існує, пропускаю: {md_file.name}")
                    continue
                md_file.rename(dest)
                log.info(f"  [cleanup] {sub_dir.parent.name}/{sub_dir.name}/"
                         f"{md_file.name} → {correct_dir.name}/{md_file.name}")
                moved_total += 1

            # Видаляємо порожню підпапку
            try:
                sub_dir.rmdir()
                log.info(f"  [cleanup] Видалено порожню папку: {sub_dir}")
            except OSError:
                log.warning(f"  [cleanup] Папка не порожня, не видалено: {sub_dir}")

    if moved_total:
        log.info(f"[cleanup] Переміщено файлів: {moved_total}")
    else:
        log.info("[cleanup] Вкладених дублів не знайдено.")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="CRE кліпінг-аналізатор [.agents/skills/cre-valuation/]"
    )
    parser.add_argument("--dry-run",       action="store_true")
    parser.add_argument("--reparse",       action="store_true")
    parser.add_argument("--file",          type=str, default=None)
    parser.add_argument("--clippings-dir", type=str, default=str(CLIPPINGS_DIR))
    args = parser.parse_args()

    if args.file:
        target = Path(args.file)
        if not target.exists():
            sys.exit(f"Файл не знайдено: {target}")
        files = [target]
    else:
        d = Path(args.clippings_dir)
        if not d.is_dir():
            sys.exit(f"Папку не знайдено: {d}")

        # Спочатку очищуємо вкладені дублі (Offices/Offices/, Warehouses/Warehouses/)
        cleanup_nested_dirs(d)

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
        r = process_file(path, dry_run=args.dry_run)
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
#  4. CLASSIFY_PROMPT:     Hotel=готель/хостел/апарт-готель
#
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    main()
