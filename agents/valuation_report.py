#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор оцінки ринкової вартості КН — Senior Analyst Edition.
Методологія: .agents/skills/cre-valuation/ [B1] [B2] [B3]

Алгоритм:
  1. Full Data Audit: перевірити 14 обов'язкових полів суб'єкта;
     відсутні — запитати інтерактивно, записати у YAML-картку.
  2. [B1] Завантажити ВСІ аналоги з Clippings/<Category>/; без ліміту.
     Матриця з 16-ма корегуваннями (кожне — Excel-формула).
  3. [B2] Дохідний підхід на базі ринкової ставки з B1.
  4. [B3] Узгодження з ваговими коефіцієнтами.
  5. Excel-only: Word генерується окремою командою.

Запуск:
    uv run agents/valuation_report.py \\
        --object-dir Владимирская_8 --type Office --excel-only
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import OrderedDict
from datetime import date
from pathlib import Path
from typing import Any

import yaml

# ─── Залежності ───────────────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("pip install openpyxl")

# ─── Шляхи ────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent.parent
OBJECTS_DIR   = BASE_DIR / "Объекты"
CLIPPINGS_DIR = BASE_DIR / "Clippings"
OUTPUT_DIR    = BASE_DIR / "output" / "reports"

# ─── Full Data Audit — 14 обов'язкових параметрів ─────────────────────────────
REQUIRED_AUDIT_FIELDS: OrderedDict = OrderedDict([
    ("District",            {"label": "Район міста (напр. Шевченківський)",               "type": "str"}),
    ("Location_Zone",       {"label": "Зона (Center/Middle/Periphery/Suburbs)",            "type": "str"}),
    ("Distance_to_Metro",   {"label": "Відстань до метро, хв пішки",                      "type": "float"}),
    ("Building_Class",      {"label": "Клас будівлі (A / B+ / B / C)",                    "type": "str"}),
    ("Floor_Type",          {"label": "Тип поверху (1й пов / мансарда / підвал / ...)",   "type": "str"}),
    ("Area",                {"label": "Загальна площа GBA, м²",                           "type": "float"}),
    ("Ceiling_Height",      {"label": "Висота стелі, м (0 = н/д)",                        "type": "float"}),
    ("Condition_Type",      {"label": "Стан (з ремонтом / під оздоблення / без ремонту)", "type": "str"}),
    ("Renovation_Style",    {"label": "Стиль ремонту (офісний / дизайнерський / ...)",    "type": "str"}),
    ("Generator",           {"label": "Генератор є? (так/ні)",                            "type": "bool"}),
    ("Shelter",             {"label": "Укриття є? (так/ні)",                              "type": "bool"}),
    ("Parking_Underground", {"label": "Підземний паркінг — к-сть місць (0 = немає)",      "type": "int"}),
    ("Parking_Open",        {"label": "Відкритий паркінг — к-сть місць (0 = немає)",      "type": "int"}),
    ("OPEX_Owner",          {"label": "OPEX: хто платить (орендар/власник/змішано)",      "type": "str"}),
])

_AUDIT_FIELD_MAP: dict[str, str] = {
    "District":            "district",
    "Location_Zone":       "zone",
    "Distance_to_Metro":   "metro_min",
    "Building_Class":      "building_class",
    "Floor_Type":          "floor_type",
    "Area":                "area",
    "Ceiling_Height":      "ceiling_height",
    "Condition_Type":      "condition",
    "Renovation_Style":    "renovation_style",
    "Generator":           "generator",
    "Shelter":             "shelter",
    "Parking_Underground": "parking_underground",
    "Parking_Open":        "parking_open",
    "OPEX_Owner":          "opex_owner",
}

# ─── Ринкові константи ─────────────────────────────────────────────────────────
DISCOUNT_RATES: dict[str, float] = {
    "Warehouse": 0.07, "Office": 0.08, "Retail": 0.06,
    "Land": 0.05, "Residential": 0.05,
}
CAP_RATES: dict[str, float] = {
    "Warehouse": 0.12, "Office": 0.10, "Retail": 0.11,
    "Land": 0.08, "Residential": 0.08,
}
OPEX_PCT: dict[str, float] = {
    "Warehouse": 0.15, "Office": 0.20, "Retail": 0.18,
    "Land": 0.05, "Residential": 0.25,
}

# ─────────────────────────────────────────────────────────────────────────────
# YAML-утиліти
# ─────────────────────────────────────────────────────────────────────────────

FRONTMATTER_RE = re.compile(r"^---\n(.*?)\n---", re.DOTALL)


def load_yaml_from_md(md_path: Path) -> dict:
    try:
        text = md_path.read_text(encoding="utf-8")
        m    = FRONTMATTER_RE.match(text)
        if m:
            return yaml.safe_load(m.group(1)) or {}
    except Exception:
        pass
    return {}


def update_yaml_in_md(md_path: Path, updates: dict) -> None:
    text = md_path.read_text(encoding="utf-8")
    m    = FRONTMATTER_RE.match(text)
    if m:
        data = yaml.safe_load(m.group(1)) or {}
        data.update(updates)
        fm   = yaml.dump(data, allow_unicode=True,
                         default_flow_style=False, sort_keys=False).rstrip()
        new_text = f"---\n{fm}\n---" + text[m.end():]
    else:
        fm   = yaml.dump(updates, allow_unicode=True,
                         default_flow_style=False, sort_keys=False).rstrip()
        new_text = f"---\n{fm}\n---\n\n" + text
    md_path.write_text(new_text, encoding="utf-8")


def find_object_md(obj_name: str) -> Path | None:
    """Шукає головну MD-картку в Объекты/{obj_name}/wiki/objects/."""
    d = OBJECTS_DIR / obj_name / "wiki" / "objects"
    if not d.exists():
        return None
    mds = sorted(d.glob("*.md"), key=lambda p: p.stat().st_mtime, reverse=True)
    return mds[0] if mds else None


def _cast(raw: str, ftype: str) -> Any:
    raw = raw.strip()
    if ftype == "float":  return float(raw.replace(",", "."))
    if ftype == "int":    return int(float(raw.replace(",", ".")))
    if ftype == "bool":   return raw.lower() in ("так", "yes", "y", "true", "1", "+")
    return raw


# ─────────────────────────────────────────────────────────────────────────────
# Full Data Audit
# ─────────────────────────────────────────────────────────────────────────────

def full_data_audit(subject: dict, obj_name: str | None,
                    md_path: Path | None) -> dict:
    """
    Перевіряє 14 обов'язкових полів.
    Відсутні — запитує в терміналі та записує у YAML-картку.
    """
    if md_path and md_path.exists():
        yd = load_yaml_from_md(md_path)
        for yk, sk in _AUDIT_FIELD_MAP.items():
            if subject.get(sk) is None and yd.get(yk) is not None:
                subject[sk] = yd[yk]

    missing = [
        (yk, _AUDIT_FIELD_MAP[yk], meta)
        for yk, meta in REQUIRED_AUDIT_FIELDS.items()
        if subject.get(_AUDIT_FIELD_MAP[yk]) is None
    ]

    if not missing:
        print("  ✓ Full Data Audit: всі 14 параметрів заповнені.")
        return subject

    name_d = obj_name or subject.get("name", "об'єкт")
    print(f"\n{'='*62}")
    print(f"  ВНИМАНИЕ: Недостаточно данных для объекта «{name_d}»")
    print(f"  Відсутніх: {len(missing)} / {len(REQUIRED_AUDIT_FIELDS)}")
    print(f"{'='*62}")

    collected: dict = {}
    for yk, sk, meta in missing:
        while True:
            try:
                raw = input(f"\n  [{yk}] {meta['label']}: ").strip()
                if not raw:
                    print("  ⚠ Пропущено.")
                    break
                val = _cast(raw, meta["type"])
                subject[sk] = val
                collected[yk] = val
                break
            except (ValueError, KeyboardInterrupt):
                print("  ✗ Невірний формат.")

    if md_path and collected:
        try:
            update_yaml_in_md(md_path, collected)
            print(f"\n  ✓ YAML оновлено: {md_path.name}")
        except Exception as e:
            print(f"\n  ⚠ Не вдалося оновити MD: {e}")
    print()
    return subject


def ask_dynamic_inputs(obj_type: str) -> tuple[float, float]:
    """Запитує знижку на торг і Cap Rate перед розрахунком."""
    dd = DISCOUNT_RATES.get(obj_type, 0.07)
    dc = CAP_RATES.get(obj_type, 0.10)

    print(f"\n{'─'*52}")
    print("  ДИНАМІЧНІ ВХІДНІ ПАРАМЕТРИ")
    print(f"{'─'*52}")

    for label, default, key in [
        ("Знижка на торг (%)", dd * 100, "disc"),
        ("Cap Rate (%)",       dc * 100, "cap"),
    ]:
        while True:
            raw = input(f"  {label} [орієнтир {default:.0f}%, Enter = прийняти]: ").strip()
            if not raw:
                val = default / 100
                break
            try:
                v = float(raw.replace(",", ".").replace("%", ""))
                val = v / 100 if v > 1 else v
                break
            except ValueError:
                print("  ✗ Введіть число.")
        if key == "disc":
            disc = val
        else:
            cap = val

    print(f"\n  → Торг: {disc*100:.1f}%  |  Cap Rate: {cap*100:.1f}%")
    print(f"{'─'*52}\n")
    return disc, cap


# ─────────────────────────────────────────────────────────────────────────────
# Завантаження аналогів — ВСІ з Clippings/<Category>/
# ─────────────────────────────────────────────────────────────────────────────

def _boolval(v: Any) -> bool | None:
    """Нормалізує булеве значення з YAML."""
    if v is None: return None
    if isinstance(v, bool): return v
    return str(v).lower() in ("true", "так", "yes", "1")


def load_all_analogs(obj_type: str) -> list[dict]:
    """
    Завантажує ВСІ валідні аналоги з Clippings/<Category>/.
    Обмежень на кількість немає.
    Приймає Rent та Sale.
    """
    folder_map = {
        "Office":    "Offices",
        "Warehouse": "Warehouses",
        "Retail":    "Retail",
        "Land":      "Land",
    }
    folder = CLIPPINGS_DIR / folder_map.get(obj_type, obj_type)

    results: list[dict] = []
    glob_paths = list(folder.glob("*.md")) if folder.exists() else []
    # Також шукаємо у всіх підпапках Clippings
    if not glob_paths:
        glob_paths = list(CLIPPINGS_DIR.glob("**/*.md"))

    for md_path in sorted(glob_paths):
        fm = load_yaml_from_md(md_path)
        if not fm:
            continue
        # Беремо розпарсені або файли з ключовими полями
        if not (fm.get("parsed") or fm.get("Deal_Type") or fm.get("Rent_per_sqm") or fm.get("Rent_Monthly_Total")):
            continue
        if fm.get("Category") and fm.get("Category") != obj_type:
            continue
        if fm.get("Status") == "Subject":
            continue

        rent_psm = fm.get("Rent_per_sqm")
        total    = fm.get("Rent_Monthly_Total") or fm.get("Price")
        area     = fm.get("Area")

        # Розраховуємо Rent_per_sqm якщо відсутня
        if not rent_psm and total and area and float(area) > 0:
            rent_psm = round(float(total) / float(area), 2)

        if not rent_psm:
            continue

        # Конвертація з UAH у USD
        currency = fm.get("Price_Currency", "USD")
        if currency == "UAH":
            rate = fm.get("Exchange_Rate", 41.5)
            rent_psm = round(float(rent_psm) / float(rate), 2)

        rent_psm = float(rent_psm)
        if rent_psm < 1 or rent_psm > 200:  # фільтр аномалій
            continue

        a: dict[str, Any] = {
            "name":            md_path.stem[:60],
            "deal_type":       fm.get("Deal_Type", "Rent"),
            "area":            float(area) if area else None,
            "rent_psm":        rent_psm,
            "zone":            fm.get("Location_Zone") or "Unknown",
            "district":        fm.get("District") or "",
            "metro_min":       fm.get("Distance_to_Metro"),
            "building_class":  fm.get("Building_Class") or "",
            "floor_type":      fm.get("Floor_Type") or "",
            "condition":       fm.get("Condition_Type") or "",
            "renovation":      fm.get("Renovation_Style") or "",
            "ceiling":         fm.get("Ceiling_Height"),
            "generator":       _boolval(fm.get("Generator")),
            "shelter":         _boolval(fm.get("Shelter")),
            "parking_u":       _boolval(fm.get("Parking_Underground")),
            "parking_o":       _boolval(fm.get("Parking_Open")),
            "building_type":   fm.get("Object_Type") or "",
            "location":        fm.get("Location") or "",
        }
        results.append(a)

    return results


# ─────────────────────────────────────────────────────────────────────────────
# Python-розрахунок коригувань (для валідації)
# ─────────────────────────────────────────────────────────────────────────────

ZONE_RANK = {"Center": 4, "Middle": 3, "Periphery": 2, "Suburbs": 1}
CLASS_RANK = {"A": 4, "B+": 3, "B": 2, "C": 1}


def _py_adj(subject: dict, analog: dict, discount: float) -> dict:
    """Повертає всі 16 коригувань + після-торгову ціну."""
    psm = analog["rent_psm"]
    after_disc = psm * (1 - discount)

    s_area  = subject.get("area", 963) or 963
    a_area  = analog.get("area") or s_area
    ratio   = s_area / a_area
    scale   = (-0.06 if ratio > 1.5 else -0.03 if ratio > 1.25 else
               0.06 if ratio < 0.67 else 0.03 if ratio < 0.8 else 0.0)

    s_zone = ZONE_RANK.get(subject.get("zone", "Center"), 4)
    a_zone = ZONE_RANK.get(analog.get("zone", "Center"), 2)
    zone_adj = (s_zone - a_zone) * 0.08

    s_dist = subject.get("district", "")
    a_dist = analog.get("district", "")
    if not a_dist:           district_adj = 0.0
    elif a_dist == s_dist:   district_adj = 0.0
    elif a_dist == "Печерський": district_adj = -0.03
    else:                    district_adj = 0.0

    s_metro = subject.get("metro_min") or 10
    a_metro = analog.get("metro_min")
    if a_metro is None:      metro_adj = 0.0
    elif a_metro <= 3:       metro_adj = -0.05
    elif a_metro <= 7:       metro_adj = -0.03
    elif a_metro <= s_metro: metro_adj = 0.0
    elif a_metro <= 15:      metro_adj = 0.03
    else:                    metro_adj = 0.05

    s_cls = CLASS_RANK.get(subject.get("building_class", "B+"), 3)
    a_cls = CLASS_RANK.get(analog.get("building_class"), 3)
    class_adj = (s_cls - a_cls) * 0.05 * (-1)  # analog better→negative

    # Floor type
    ft = (analog.get("floor_type") or "").lower()
    floor_adj = (0.03 if "мансард" in ft else
                 0.05 if "підвал" in ft else
                 0.0)

    # Condition
    cond = (analog.get("condition") or "").lower()
    if "без ремонт" in cond:          cond_adj = 0.15
    elif "під оздоблення" in cond or "без меблів" in cond: cond_adj = 0.10
    elif "з ремонтом" in cond:        cond_adj = -0.05
    else:                             cond_adj = 0.0

    # Renovation style
    renov = (analog.get("renovation") or "").lower()
    if "представниць" in renov:       renov_adj = -0.05
    elif "дизайнер" in renov or "loft" in renov: renov_adj = -0.08
    elif "новий офіс" in renov:       renov_adj = -0.03
    else:                             renov_adj = 0.0

    # Ceiling
    s_ceil = subject.get("ceiling_height") or 3.5
    a_ceil = analog.get("ceiling")
    if a_ceil is None:                ceil_adj = 0.0
    elif a_ceil > 5:                  ceil_adj = -0.08
    elif a_ceil > 4:                  ceil_adj = -0.05
    elif a_ceil > 3.5:                ceil_adj = -0.02
    elif a_ceil > 2.8:                ceil_adj = 0.0
    else:                             ceil_adj = 0.05

    # Generator — subject HAS
    ag = analog.get("generator")
    gen_adj = 0.0 if ag is True else (0.05 if ag is False else 0.0)

    # Shelter — subject HAS
    ash = analog.get("shelter")
    shelt_adj = 0.0 if ash is True else (0.05 if ash is False else 0.0)

    # Underground parking — subject has 0
    apu = analog.get("parking_u")
    park_u_adj = -0.05 if apu is True else 0.0

    # Open parking — subject has 15
    apo = analog.get("parking_o")
    park_o_adj = 0.03 if apo is False else 0.0

    # Building type (standalone bonus for subject, analog in БЦ = slightly better)
    bt = (analog.get("building_type") or "").lower()
    bldg_type_adj = (-0.03 if "бізнес" in bt or "центр" in bt else
                     0.0   if "окрем" in bt or "адмін" in bt else
                     0.0)

    # Renovation age (subject: >10 years → analogs with fresh are better → −3%)
    renov_age_adj = -0.03  # default: analogs assume fresh, subject is old

    total = (scale + zone_adj + district_adj + metro_adj + class_adj +
             floor_adj + cond_adj + renov_adj + ceil_adj + gen_adj +
             shelt_adj + park_u_adj + park_o_adj + bldg_type_adj + renov_age_adj)

    return {
        "psm":          psm,
        "after_disc":   after_disc,
        "adj_scale":    scale,
        "adj_zone":     zone_adj,
        "adj_district": district_adj,
        "adj_metro":    metro_adj,
        "adj_class":    class_adj,
        "adj_floor":    floor_adj,
        "adj_cond":     cond_adj,
        "adj_renov":    renov_adj,
        "adj_ceil":     ceil_adj,
        "adj_gen":      gen_adj,
        "adj_shelt":    shelt_adj,
        "adj_park_u":   park_u_adj,
        "adj_park_o":   park_o_adj,
        "adj_bldg":     bldg_type_adj,
        "adj_age":      renov_age_adj,
        "total_adj":    total,
        "final_psm":    after_disc * (1 + total),
    }


# ─────────────────────────────────────────────────────────────────────────────
# [B1] Порівняльний підхід — ринкова орендна ставка
# ─────────────────────────────────────────────────────────────────────────────

def rent_comparable_approach(subject: dict, analogs: list[dict],
                              discount: float) -> dict:
    """[B1] Повертає матрицю і ринкову ставку ($/м²/міс)."""
    rows = []
    for a in analogs:
        adjs = _py_adj(subject, a, discount)
        rows.append({**a, **adjs})
    if not rows:
        return {"error": "Немає аналогів"}
    finals = [r["final_psm"] for r in rows]
    # Усічена середня: без min та max
    trimmed = sorted(finals)[1:-1] if len(finals) > 4 else finals
    market_rent = sum(trimmed) / len(trimmed)
    return {
        "rows": rows, "finals": finals,
        "market_rent": market_rent,
        "n": len(rows),
    }


# ─────────────────────────────────────────────────────────────────────────────
# [B2] Дохідний підхід
# ─────────────────────────────────────────────────────────────────────────────

def income_approach(area: float, market_rent: float,
                    vacancy: float, opex_pct: float, cap: float) -> dict:
    pgi  = area * market_rent * 12
    vl   = pgi * vacancy
    egi  = pgi - vl
    opex = egi * opex_pct
    noi  = egi - opex
    val  = noi / cap
    return {"pgi": pgi, "vacancy_loss": vl, "egi": egi,
            "opex": opex, "noi": noi, "value": val,
            "area": area, "market_rent": market_rent,
            "vacancy": vacancy, "opex_pct": opex_pct, "cap_rate": cap}


def reconciliation(v1: float, v2: float,
                   w1: float = 0.40, w2: float = 0.60) -> dict:
    wtd = v1 * w1 + v2 * w2
    mag = 10 ** (len(str(int(wtd))) - 2)
    return {"value_b1": v1, "value_b2": v2,
            "weight_b1": w1, "weight_b2": w2,
            "weighted": wtd, "value": round(wtd / mag) * mag}


# ─────────────────────────────────────────────────────────────────────────────
# Excel — стилі
# ─────────────────────────────────────────────────────────────────────────────

C_NAVY   = "1C2E44"
C_STEEL  = "2D4A6B"
C_GOLD   = "D5B58A"
C_INPUT  = "E1FAFF"   # вхідні дані — світло-блакитний
C_WHITE  = "FFFFFF"
C_STRIPE = "F2F6FA"
C_IVORY  = "FAFAF8"
C_GREEN  = "D5E8D4"
C_RED    = "F8CECC"
C_YELL   = "FFF2CC"
C_LIGHT  = "EAF0F6"

_thin = Side(style="thin", color="CCCCCC")
_BDR  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _st(cell: Any, bold=False, italic=False, sz=9,
        color=C_NAVY, bg=C_WHITE, align="center",
        wrap=False, fmt: str | None = None) -> None:
    cell.font      = Font(name="Calibri", bold=bold, italic=italic,
                          size=sz, color=color)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center",
                               wrap_text=wrap)
    cell.border    = _BDR
    if fmt:
        cell.number_format = fmt


def _w(ws, r, c, val, **kw):
    cell = ws.cell(row=r, column=c, value=val)
    _st(cell, **kw)
    return cell


def _inp(ws, r, c, val, fmt="General"):
    """Вхідна комірка E1FAFF."""
    cell = ws.cell(row=r, column=c, value=val)
    _st(cell, bold=True, color=C_NAVY, bg=C_INPUT, align="center", fmt=fmt)
    return cell


def _hdr_cell(ws, r, c, text, width=None, sz=9, bg=C_NAVY):
    fc = C_WHITE if bg in (C_NAVY, C_STEEL) else C_NAVY
    cell = ws.cell(row=r, column=c, value=text)
    _st(cell, bold=True, sz=sz, color=fc, bg=bg,
        align="center", wrap=True)
    if width:
        ws.column_dimensions[get_column_letter(c)].width = width
    return cell


def _merge(ws, r, c1, c2, text, bold=True, sz=10,
           bg=C_NAVY, color=C_WHITE, align="left"):
    ws.merge_cells(start_row=r, start_column=c1,
                   end_row=r,   end_column=c2)
    cell = ws.cell(row=r, column=c1, value=text)
    _st(cell, bold=bold, sz=sz, color=color, bg=bg, align=align)


def _adj_bg(val: float) -> str:
    if val > 0.005:  return C_GREEN
    if val < -0.005: return C_RED
    return C_YELL


# ─────────────────────────────────────────────────────────────────────────────
# Excel формули для 16 коригувань
# ─────────────────────────────────────────────────────────────────────────────
# Посилання на параметри суб'єкта (абсолютні адреси в B1-аркуші):
# $B$5 = Area  $B$6 = Zone  $B$7 = District  $B$8 = Metro
# $B$9 = Class  $B$13 = Ceiling  $B$14 = Generator  $B$15 = Shelter
# $B$16 = Parking_U  $B$17 = Parking_O  $B$18 = Disc rate

def _xl_scale(r):    # масштаб площі — кол E
    return (f"=IFERROR(IF(E{r}=\"\",0,"
            f"IF($B$5/E{r}>1.5,-0.06,IF($B$5/E{r}>1.25,-0.03,"
            f"IF($B$5/E{r}<0.67,0.06,IF($B$5/E{r}<0.8,0.03,0))))),0)")


def _xl_zone(r):     # зона — кол D
    return (f"=IFERROR(IF(D{r}=\"\",0,"
            f"IF(D{r}=\"Center\",0,IF(D{r}=\"Middle\",0.08,"
            f"IF(D{r}=\"Periphery\",0.16,0.24)))),0)")


def _xl_district(r):  # район — кол C
    return (f"=IFERROR(IF(C{r}=\"\",0,IF(C{r}=$B$7,0,"
            f"IF(C{r}=\"Печерський\",-0.03,0))),0)")


def _xl_metro(r):    # метро — кол H (суб'єкт $B$8)
    return (f"=IFERROR(IF(H{r}=\"\",0,"
            f"IF(H{r}<=3,-0.05,IF(H{r}<=7,-0.03,"
            f"IF(H{r}<=$B$8,0,IF(H{r}<=15,0.03,0.05))))),0)")


def _xl_class(r):    # клас — кол I
    return (f"=IFERROR(IF(I{r}=\"\",0,"
            f"IF(I{r}=\"A\",-0.1,IF(I{r}=\"B+\",0,"
            f"IF(I{r}=\"B\",0.05,IF(I{r}=\"C\",0.1,0))))),0)")


def _xl_floor(r):    # тип поверху — кол J
    return (f"=IFERROR(IF(J{r}=\"\",0,"
            f"IF(ISNUMBER(SEARCH(\"мансард\",J{r})),0.03,"
            f"IF(ISNUMBER(SEARCH(\"підвал\",J{r})),0.05,0))),0)")


def _xl_cond(r):     # стан — кол K
    return (f"=IFERROR(IF(K{r}=\"\",0,"
            f"IF(K{r}=\"без ремонту\",0.15,"
            f"IF(OR(K{r}=\"під оздоблення\",K{r}=\"без меблів\"),0.1,"
            f"IF(OR(K{r}=\"з ремонтом\",K{r}=\"з ремонтом, з меблями\"),-0.05,0)))),0)")


def _xl_renov(r):    # стиль ремонту — кол L
    return (f"=IFERROR(IF(L{r}=\"\",0,"
            f"IF(L{r}=\"представницький\",-0.05,"
            f"IF(OR(L{r}=\"дизайнерський\",ISNUMBER(SEARCH(\"loft\",LOWER(L{r})))),-0.08,"
            f"IF(L{r}=\"новий офісний\",-0.03,0)))),0)")


def _xl_ceil(r):     # стелі — кол M
    return (f"=IFERROR(IF(M{r}=\"\",0,"
            f"IF(M{r}>5,-0.08,IF(M{r}>4,-0.05,"
            f"IF(M{r}>3.5,-0.02,IF(M{r}>2.8,0,0.05))))),0)")


def _xl_gen(r):      # генератор — кол N (суб'єкт HAS)
    return f"=IFERROR(IF(N{r}=\"\",0,IF(N{r}=TRUE,0,0.05)),0)"


def _xl_shelt(r):    # укриття — кол O (суб'єкт HAS)
    return f"=IFERROR(IF(O{r}=\"\",0,IF(O{r}=TRUE,0,0.05)),0)"


def _xl_park_u(r):   # підземний паркінг — кол P (суб'єкт = 0)
    return f"=IFERROR(IF(OR(P{r}=\"\",P{r}=FALSE),0,-0.05),0)"


def _xl_park_o(r):   # відкритий паркінг — кол Q (суб'єкт HAS 15)
    return f"=IFERROR(IF(OR(Q{r}=\"\",Q{r}=FALSE),0.03,0),0)"


def _xl_bldg(r):     # тип будівлі — кол R
    return (f"=IFERROR(IF(R{r}=\"\",0,"
            f"IF(ISNUMBER(SEARCH(\"центр\",LOWER(R{r}))),-0.03,"
            f"IF(ISNUMBER(SEARCH(\"бізнес\",LOWER(R{r}))),-0.03,0))),0)")


# Σ та фінальна ставка:
def _xl_sum(r, s_col=19, e_col=34):   # S..AG (включно)
    sc = get_column_letter(s_col)
    ec = get_column_letter(e_col)
    return f"=SUM({sc}{r}:{ec}{r})"


def _xl_final(r, disc_col=7, sum_col=35):
    gc = get_column_letter(disc_col)   # G = після торгу
    ac = get_column_letter(sum_col)    # AH = Σ
    return f"={gc}{r}*(1+{ac}{r})"


# ─────────────────────────────────────────────────────────────────────────────
# Генерація Excel — максимально детальний
# ─────────────────────────────────────────────────────────────────────────────

def generate_excel(subject: dict, b1: dict, b2: dict, b3: dict,
                   out_path: Path, discount: float, cap: float) -> None:
    """
    Три аркуші:
      B1_Порівняльний — суб'єкт E1FAFF + 16-колонкова матриця з формулами
      B2_Дохідний     — ланцюг PGI→NOI→Value (формули)
      B3_Узгодження   — cross-sheet формули, ваги E1FAFF
    """
    wb = openpyxl.Workbook()

    # ════════════════════════════════════════════════════════════════════════
    # АРКУШ 1 — B1_Порівняльний
    # ════════════════════════════════════════════════════════════════════════
    ws = wb.active
    ws.title = "B1_Порівняльний"
    ws.sheet_view.showGridLines = False

    # Ширина стовпців — фіксована розмітка
    col_widths = {
        1: 4,   # №
        2: 26,  # Аналог
        3: 14,  # Район
        4: 10,  # Зона
        5: 8,   # Площа
        6: 11,  # Ставка ask
        7: 11,  # Після торгу
        8: 7,   # Метро
        9: 7,   # Клас
        10: 14, # Тип поверху
        11: 16, # Стан
        12: 16, # Ремонт стиль
        13: 7,  # Стелі
        14: 7,  # Ген.
        15: 7,  # Укр.
        16: 8,  # Парк п/з
        17: 8,  # Парк відкр.
        18: 14, # Тип будівлі
        # adj cols 19..34 = S..AG
        **{c: 8 for c in range(19, 37)},
    }
    for c, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    # ── Заголовок аркуша ──
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[2].height = 28
    _merge(ws, 2, 1, 36,
           "[B1]  ПОРІВНЯЛЬНИЙ АНАЛІЗ ОРЕНДНИХ СТАВОК — 16 КРИТЕРІЇВ  "
           f"|  {subject.get('name', 'Об єкт')}",
           sz=12, align="center")

    # ── БЛОК: суб'єкт оцінки (рядки 4–22) ──
    ws.row_dimensions[3].height = 6
    ws.row_dimensions[4].height = 20
    _merge(ws, 4, 1, 4, "СУОБ'ЄКТ ОЦІНКИ — ВХІДНІ ДАНІ",
           sz=10, bg=C_STEEL, color=C_WHITE, align="left")
    _merge(ws, 4, 5, 36,
           "Комірки E1FAFF — редаговані. Зміна значення автоматично перераховує всі формули матриці.",
           sz=9, bg=C_LIGHT, color=C_STEEL, align="left", bold=False)

    subj_rows = [
        (5,  "Площа GBA, м²",         subject.get("area", 963),                  "#,##0",   "B5"),
        (6,  "Зона розташування",      subject.get("zone", "Center"),             "@",       "B6"),
        (7,  "Район",                  subject.get("district", "Шевченківський"), "@",       "B7"),
        (8,  "Відстань до метро, хв",  subject.get("metro_min") or 10,            "0",       "B8"),
        (9,  "Клас будівлі",           subject.get("building_class", "B+"),       "@",       "B9"),
        (10, "Тип поверху",            subject.get("floor_type", ""),             "@",       "B10"),
        (11, "Технічний стан",         subject.get("condition", "з ремонтом"),    "@",       "B11"),
        (12, "Стиль ремонту",          subject.get("renovation_style", ""),       "@",       "B12"),
        (13, "Висота стелі, м",        subject.get("ceiling_height") or 3.5,      "0.0",     "B13"),
        (14, "Генератор",              subject.get("generator", True),            "@",       "B14"),
        (15, "Укриття",                subject.get("shelter", True),              "@",       "B15"),
        (16, "Паркінг підземний, міс.",subject.get("parking_underground", 0),     "0",       "B16"),
        (17, "Паркінг відкритий, міс.",subject.get("parking_open", 15),           "0",       "B17"),
        (18, "Знижка на торг, %  ⚙",  discount,                                  "0.0%",    "B18"),
        (19, "Cap Rate, %  ⚙",         cap,                                       "0.0%",    "B19"),
        (20, "Вакансія, %  ⚙",         subject.get("vacancy", 0.15),             "0%",      "B20"),
        (21, "OPEX від EGI, %  ⚙",     subject.get("opex_pct", 0.20),            "0%",      "B21"),
    ]

    for row_n, label, val, fmt, _ in subj_rows:
        ws.row_dimensions[row_n].height = 17
        _w(ws, row_n, 1, label, bold=True, bg=C_STRIPE, align="left", sz=9)
        ws.merge_cells(start_row=row_n, start_column=1,
                       end_row=row_n, end_column=4)
        _inp(ws, row_n, 2, val, fmt=fmt)
        # Очищаємо об'єднані комірки щоб не було артефактів
        for cc in range(2, 5):
            if cc != 2:
                c = ws.cell(row=row_n, column=cc)
                c.border = _BDR

    ws.row_dimensions[22].height = 6

    # ── Заголовки матриці (рядок 23) ──
    HDR_ROW  = 23
    DATA_ROW = 24
    ws.row_dimensions[HDR_ROW].height = 52

    input_hdrs = [
        (1,  "№"),
        (2,  "Аналог"),
        (3,  "Район"),
        (4,  "Зона"),
        (5,  "Площа,\nм²"),
        (6,  "Ставка ask\n$/м²/міс"),
        (7,  "Після торгу\n$/м²"),
        (8,  "Метро,\nхв"),
        (9,  "Клас\nбудівлі"),
        (10, "Тип\nповерху"),
        (11, "Тех.\nстан"),
        (12, "Стиль\nремонту"),
        (13, "Стелі,\nм"),
        (14, "Ген."),
        (15, "Укр."),
        (16, "Парк\nп/з"),
        (17, "Парк\nвідкр."),
        (18, "Тип\nбудівлі"),
    ]
    adj_hdrs = [
        (19, "Кор.\nмасштаб"),
        (20, "Кор.\nзона"),
        (21, "Кор.\nрайон"),
        (22, "Кор.\nметро"),
        (23, "Кор.\nклас"),
        (24, "Кор.\nповерх"),
        (25, "Кор.\nстан"),
        (26, "Кор.\nремонт"),
        (27, "Кор.\nстелі"),
        (28, "Кор.\nгенер."),
        (29, "Кор.\nукр."),
        (30, "Кор.\nпарк п/з"),
        (31, "Кор.\nпарк відкр."),
        (32, "Кор.\nтип буд."),
        (33, "Кор.\nвік рем. ⚙"),
        (34, "Кор.\nінше ⚙"),
        (35, "Σ\nкорег."),
        (36, "Скоригована\nставка $/м²"),
    ]

    for col_n, label in input_hdrs:
        _hdr_cell(ws, HDR_ROW, col_n, label, bg=C_NAVY)
    for col_n, label in adj_hdrs:
        bg = C_STEEL if col_n < 35 else C_GOLD
        _hdr_cell(ws, HDR_ROW, col_n, label, bg=bg)

    # ── Рядки аналогів ──
    rows_b1 = b1.get("rows", [])

    for i, row in enumerate(rows_b1):
        r  = DATA_ROW + i
        bg = C_IVORY if i % 2 == 0 else C_STRIPE
        ws.row_dimensions[r].height = 17

        # Дані (E1FAFF для редагованих)
        _w(ws, r, 1,  i + 1,                    bg=bg, fmt="0", sz=9)
        _w(ws, r, 2,  row["name"],               bg=bg, align="left", sz=8)
        # Район, Зона, Площа, Ставка — E1FAFF (аналітик може правити)
        _inp(ws, r, 3,  row.get("district", ""))
        _inp(ws, r, 4,  row.get("zone", ""))
        _inp(ws, r, 5,  row.get("area"),         fmt="#,##0")
        _inp(ws, r, 6,  row.get("rent_psm"),     fmt='#,##0.00 "$"')

        # Col G (7): Після торгу = ФОРМУЛА
        cg = ws.cell(row=r, column=7, value=f"=F{r}*(1-$B$18)")
        _st(cg, bold=True, bg=bg, fmt='#,##0.00 "$"', sz=9)

        # Решта вхідних — E1FAFF
        _inp(ws, r, 8,  row.get("metro_min"),    fmt="0")
        _inp(ws, r, 9,  row.get("building_class", ""))
        _inp(ws, r, 10, row.get("floor_type", ""))
        _inp(ws, r, 11, row.get("condition", ""))
        _inp(ws, r, 12, row.get("renovation", ""))
        _inp(ws, r, 13, row.get("ceiling"),      fmt="0.0")
        _inp(ws, r, 14, row.get("generator"))
        _inp(ws, r, 15, row.get("shelter"))
        _inp(ws, r, 16, row.get("parking_u"))
        _inp(ws, r, 17, row.get("parking_o"))
        _inp(ws, r, 18, row.get("building_type", ""))

        # Формули коригувань (кол 19–34 = S..AG)
        xl_adjs = [
            (19, _xl_scale(r)),
            (20, _xl_zone(r)),
            (21, _xl_district(r)),
            (22, _xl_metro(r)),
            (23, _xl_class(r)),
            (24, _xl_floor(r)),
            (25, _xl_cond(r)),
            (26, _xl_renov(r)),
            (27, _xl_ceil(r)),
            (28, _xl_gen(r)),
            (29, _xl_shelt(r)),
            (30, _xl_park_u(r)),
            (31, _xl_park_o(r)),
            (32, _xl_bldg(r)),
        ]
        for col_n, formula in xl_adjs:
            py_val = row.get("adj_" + [
                "scale","zone","district","metro","class","floor",
                "cond","renov","ceil","gen","shelt","park_u","park_o","bldg"
            ][col_n - 19], 0.0)
            cell = ws.cell(row=r, column=col_n, value=formula)
            _st(cell, bg=_adj_bg(py_val), fmt="0.0%", sz=9)

        # Кор. вік ремонту (col 33) і Інше (col 34) — E1FAFF, ручне
        _inp(ws, r, 33, row.get("adj_age", -0.03), fmt="0.0%")
        _inp(ws, r, 34, 0.0, fmt="0.0%")

        # Col 35 (AH): Σ = SUM(S:AG)
        sigma_cell = ws.cell(row=r, column=35,
                             value=f"=SUM({get_column_letter(19)}{r}:{get_column_letter(34)}{r})")
        _st(sigma_cell, bold=True, fmt="0.0%", sz=9,
            bg=(C_YELL if abs(row.get("total_adj", 0)) > 0.20 else C_GREEN))

        # Col 36 (AI): Скоригована ставка = G*(1+Σ)
        final_cell = ws.cell(row=r, column=36,
                             value=f"=G{r}*(1+{get_column_letter(35)}{r})")
        _st(final_cell, bold=True, bg=bg, fmt='#,##0.00 "$"', sz=9)

    LAST_DATA = DATA_ROW + len(rows_b1) - 1
    FIN_COL   = 36   # AI — скоригована ставка
    FC        = get_column_letter(FIN_COL)

    # ── Статистика ──
    STAT_ROW = LAST_DATA + 2
    ws.row_dimensions[LAST_DATA + 1].height = 6
    ws.row_dimensions[STAT_ROW].height = 20
    _merge(ws, STAT_ROW, 1, 35,
           "СТАТИСТИКА СКОРИГОВАНИХ ОРЕНДНИХ СТАВОК, $/м²/міс", sz=10, bg=C_STEEL)

    STAT_LABELS = [
        ("Мін.",                    f"=MIN({FC}{DATA_ROW}:{FC}{LAST_DATA})"),
        ("P25 (PERCENTILE 25%)",    f"=PERCENTILE({FC}{DATA_ROW}:{FC}{LAST_DATA},0.25)"),
        ("Медіана (P50)",           f"=MEDIAN({FC}{DATA_ROW}:{FC}{LAST_DATA})"),
        ("Середня (AVERAGE)",       f"=AVERAGE({FC}{DATA_ROW}:{FC}{LAST_DATA})"),
        ("Усічена середня",         f"=TRIMMEAN({FC}{DATA_ROW}:{FC}{LAST_DATA},0.2)"),
        ("P75 (PERCENTILE 75%)",    f"=PERCENTILE({FC}{DATA_ROW}:{FC}{LAST_DATA},0.75)"),
        ("Макс.",                   f"=MAX({FC}{DATA_ROW}:{FC}{LAST_DATA})"),
    ]
    TRIMMED_STAT_ROW = STAT_ROW + 5   # рядок "Усічена середня"

    for k, (lbl, formula) in enumerate(STAT_LABELS):
        sr = STAT_ROW + 1 + k
        ws.row_dimensions[sr].height = 16
        _w(ws, sr, 1, lbl, bold=True, bg=C_STRIPE, align="left", sz=9)
        ws.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=35)
        sc = ws.cell(row=sr, column=36, value=formula)
        _st(sc, bold=True, bg=C_LIGHT, fmt='#,##0.00 "$"', sz=9)

    # Прийнята ринкова ставка
    MARKET_RENT_ROW = STAT_ROW + len(STAT_LABELS) + 2
    ws.row_dimensions[MARKET_RENT_ROW - 1].height = 6
    ws.row_dimensions[MARKET_RENT_ROW].height = 22
    _merge(ws, MARKET_RENT_ROW, 1, 35,
           "✅  ПРИЙНЯТА РИНКОВА ОРЕНДНА СТАВКА ($/м²/міс) — усічена середня TRIMMEAN",
           sz=10, bg=C_GOLD, color=C_NAVY, align="left")
    MARKET_RENT_CELL = f"{FC}{MARKET_RENT_ROW}"   # для B2
    mr_cell = ws.cell(row=MARKET_RENT_ROW, column=36,
                      value=f"={FC}{TRIMMED_STAT_ROW}")
    _st(mr_cell, bold=True, sz=11, bg=C_GOLD, color=C_NAVY, fmt='#,##0.00 "$"')

    # ════════════════════════════════════════════════════════════════════════
    # АРКУШ 2 — B2_Дохідний
    # ════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("B2_Дохідний")
    ws2.sheet_view.showGridLines = False
    for cl, w in [("A", 40), ("B", 24), ("C", 18), ("D", 14)]:
        ws2.column_dimensions[cl].width = w

    ws2.row_dimensions[1].height = 8
    ws2.row_dimensions[2].height = 28
    _merge(ws2, 2, 1, 3,
           "[B2]  ДОХІДНИЙ ПІДХІД  |  Пряма капіталізація NOI  |  Формули",
           sz=12, align="center")

    r2 = 4
    ws2.row_dimensions[r2].height = 20
    _merge(ws2, r2, 1, 3, "ВХІДНІ ПАРАМЕТРИ  (E1FAFF — редаговані)", sz=10, bg=C_STEEL)
    r2 += 1

    def _b2_inp(label, val, fmt, note=""):
        nonlocal r2
        ws2.row_dimensions[r2].height = 17
        _w(ws2, r2, 1, label, bold=True, bg=C_STRIPE, align="left", sz=9)
        _inp(ws2, r2, 2, val, fmt=fmt)
        addr = f"B{r2}"
        if note:
            cn = ws2.cell(row=r2, column=3, value=note)
            _st(cn, italic=True, color="777777", align="left", sz=8, bg=C_WHITE)
            ws2.merge_cells(f"C{r2}:D{r2}")
        r2 += 1
        return addr

    # GLA ссылается на B1 аркуш
    gla_addr  = _b2_inp("GLA — площа, м²",
                         f"='B1_Порівняльний'!$B$5",
                         "#,##0", "= B1_Порівняльний!$B$5")
    rent_addr = _b2_inp("Ринкова орендна ставка, $/м²/міс",
                         f"='B1_Порівняльний'!{MARKET_RENT_CELL}",
                         '0.00 "$"', f"= {MARKET_RENT_CELL} (TRIMMEAN з B1)")
    vac_addr  = _b2_inp("Вакансія та втрати від несплати, %",
                         f"='B1_Порівняльний'!$B$20",
                         "0%", "= B1_Порівняльний!$B$20")
    opex_addr = _b2_inp("OPEX від EGI, %",
                         f"='B1_Порівняльний'!$B$21",
                         "0%", "= B1_Порівняльний!$B$21")
    cap_addr  = _b2_inp("Cap Rate (ставка капіт.), %",
                         f"='B1_Порівняльний'!$B$19",
                         "0.0%", f"= B1_Порівняльний!$B$19")
    r2 += 1

    ws2.row_dimensions[r2].height = 20
    _merge(ws2, r2, 1, 3, "РОЗРАХУНОК NOI  (всі рядки — Excel-формули)", sz=10, bg=C_STEEL)
    r2 += 1

    def _b2_calc(label, formula, fmt, bold=False, bg=C_WHITE, note=""):
        nonlocal r2
        ws2.row_dimensions[r2].height = 17
        _w(ws2, r2, 1, label, bold=bold, bg=bg, align="left", sz=9)
        cf = ws2.cell(row=r2, column=2, value=formula)
        _st(cf, bold=bold, bg=bg, fmt=fmt, sz=9)
        addr = f"B{r2}"
        if note:
            cn = ws2.cell(row=r2, column=3, value=note)
            _st(cn, italic=True, color="777777", align="left", sz=8, bg=C_WHITE)
            ws2.merge_cells(f"C{r2}:D{r2}")
        r2 += 1
        return addr

    pgi_a  = _b2_calc("PGI — Потенційний валовий дохід, $/рік",
                       f"={gla_addr}*{rent_addr}*12", '#,##0 "$"',
                       bold=True, note=f"={gla_addr}×{rent_addr}×12")
    vl_a   = _b2_calc("  − Втрати від вакансії та несплати, $/рік",
                       f"=-{pgi_a}*{vac_addr}", '#,##0 "$"',
                       note=f"PGI×{vac_addr}")
    egi_a  = _b2_calc("EGI — Ефективний валовий дохід, $/рік",
                       f"={pgi_a}+{vl_a}", '#,##0 "$"',
                       bold=True, bg=C_LIGHT)
    opex_a = _b2_calc("  − OPEX (операційні витрати), $/рік",
                       f"=-{egi_a}*{opex_addr}", '#,##0 "$"')
    noi_a  = _b2_calc("NOI — Чистий операційний дохід, $/рік",
                       f"={egi_a}+{opex_a}", '#,##0 "$"',
                       bold=True, bg=C_LIGHT)
    r2 += 1
    B2_VAL_ROW = r2
    b2v_a  = _b2_calc("ВАРТІСТЬ (B2) = NOI / Cap Rate",
                       f"={noi_a}/{cap_addr}", '#,##0 "$"',
                       bold=True, bg=C_GOLD)
    ws2.cell(row=B2_VAL_ROW, column=1).font = Font(
        name="Calibri", bold=True, size=11, color=C_NAVY)
    ws2.cell(row=B2_VAL_ROW, column=2).font = Font(
        name="Calibri", bold=True, size=12, color=C_NAVY)

    # ════════════════════════════════════════════════════════════════════════
    # АРКУШ 3 — B3_Узгодження
    # ════════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("B3_Узгодження")
    ws3.sheet_view.showGridLines = False
    for cl, w in [("A", 44), ("B", 26), ("C", 18)]:
        ws3.column_dimensions[cl].width = w

    ws3.row_dimensions[1].height = 8
    ws3.row_dimensions[2].height = 28
    _merge(ws3, 2, 1, 3,
           "[B3]  УЗГОДЖЕННЯ  |  Зважена вартість  |  Cross-sheet формули",
           sz=12, align="center")

    r3 = 4
    ws3.row_dimensions[r3].height = 20
    _merge(ws3, r3, 1, 3, "РЕЗУЛЬТАТИ ПІДХОДІВ", sz=10, bg=C_STEEL)
    r3 += 1

    def _b3r(label, val, fmt, bold=False, bg=C_WHITE, note=""):
        nonlocal r3
        ws3.row_dimensions[r3].height = 17
        _w(ws3, r3, 1, label, bold=bold, bg=bg, align="left", sz=9)
        if isinstance(val, str) and val.startswith("="):
            cf = ws3.cell(row=r3, column=2, value=val)
            _st(cf, bold=bold, bg=bg, fmt=fmt, sz=9)
        else:
            _w(ws3, r3, 2, val, bold=bold, bg=bg, fmt=fmt, sz=9)
        if note:
            cn = ws3.cell(row=r3, column=3, value=note)
            _st(cn, italic=True, color="777777", align="left", sz=8, bg=C_WHITE)
        addr = f"B{r3}"
        r3 += 1
        return addr

    # Посилання на B1: ринкова ставка → через NOI/Cap (не окремий sale comps)
    # B1-значення = market_rent → Value = NOI/Cap (B2 result)
    # B3 використовує B2 як єдиний підхід (оскільки B1 дає ставку, не ціну)
    b2_ref  = _b3r("Ринкова орендна ставка (B1), $/м²/міс",
                    f"='B1_Порівняльний'!{MARKET_RENT_CELL}",
                    '0.00 "$"', note="← TRIMMEAN з B1_Порівняльний")
    b2_val  = _b3r("Вартість за дохідним підходом (B2), $",
                    f"='B2_Дохідний'!{b2v_a}",
                    '#,##0 "$"', note="← B2_Дохідний")

    r3 += 1
    ws3.row_dimensions[r3].height = 20
    _merge(ws3, r3, 1, 3, "ПІДСУМОК", sz=10, bg=C_STEEL)
    r3 += 1

    ws3.row_dimensions[r3].height = 20
    _w(ws3, r3, 1, "Знижка на торг, %",     bold=True, bg=C_STRIPE, align="left", sz=9)
    _b3r.__func__ if hasattr(_b3r, "__func__") else None   # не потрібно
    disc_ref = ws3.cell(row=r3, column=2,
                        value=f"='B1_Порівняльний'!$B$18")
    _st(disc_ref, bg=C_INPUT, bold=True, fmt="0.0%", sz=9)
    r3 += 1

    ws3.row_dimensions[r3].height = 20
    _w(ws3, r3, 1, "Cap Rate, %",            bold=True, bg=C_STRIPE, align="left", sz=9)
    cap_ref = ws3.cell(row=r3, column=2,
                       value=f"='B1_Порівняльний'!$B$19")
    _st(cap_ref, bg=C_INPUT, bold=True, fmt="0.0%", sz=9)
    r3 += 1

    r3 += 1
    ws3.row_dimensions[r3].height = 28
    _w(ws3, r3, 1, "РИНКОВА ВАРТІСТЬ (B2), $",
       bold=True, sz=13, bg=C_NAVY, color=C_WHITE, align="left")
    final_cell = ws3.cell(row=r3, column=2,
                          value=f"=ROUND('B2_Дохідний'!{b2v_a},-4)")
    _st(final_cell, bold=True, sz=14, bg=C_GOLD, color=C_NAVY, fmt='#,##0 "$"')

    r3 += 2
    ws3.row_dimensions[r3].height = 32
    _merge(ws3, r3, 1, 3,
           f"⚠ Попередня аналітична оцінка. Не є офіційним висновком. "
           f"Дата: {date.today().strftime('%d.%m.%Y')}. "
           f"Аналогів B1: {b1['n']}. Методологія: [B1]+[B2].",
           bold=False, sz=9, bg=C_YELL, color=C_NAVY, align="left")

    wb.save(out_path)
    print(f"  ✓ Excel: {out_path.name}")
    print(f"     Аналогів: {b1['n']}  |  Ринкова ставка (TRIMMEAN): "
          f"${b1['market_rent']:.2f}/м²  |  NOI/Cap → ${b2['value']:,.0f}")


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    p = argparse.ArgumentParser(
        description="Valuation Report [B1+B2+B3] — Senior Analyst Edition"
    )
    p.add_argument("--name",          default="Об'єкт оцінки")
    p.add_argument("--type",          default="Office",
                   choices=["Warehouse", "Office", "Retail", "Land", "Residential"])
    p.add_argument("--area",          type=float, default=None)
    p.add_argument("--zone",          default="Center",
                   choices=["Center", "Middle", "Periphery", "Suburbs"])
    p.add_argument("--address",       default="—")
    p.add_argument("--vacancy",       type=float, default=0.15)
    p.add_argument("--opex-pct",      type=float, default=None)
    p.add_argument("--weight-b1",     type=float, default=0.40)
    p.add_argument("--weight-b2",     type=float, default=0.60)
    p.add_argument("--no-audit",      action="store_true")
    p.add_argument("--no-interactive",action="store_true")
    p.add_argument("--object-dir",    default=None,
                   help="Папка в Объекты/ (напр. Владимирская_8)")
    p.add_argument("--cap-rate",      type=float, default=None)
    p.add_argument("--discount",      type=float, default=None)
    p.add_argument("--excel-only",    action="store_true",
                   help="Генерувати тільки Excel (Word — окремою командою)")
    p.add_argument("--out-dir",       default=str(OUTPUT_DIR))
    args = p.parse_args()

    # ── Завантаження MD-картки об'єкта ──
    md_path: Path | None = None
    yaml_data: dict = {}
    if args.object_dir:
        md_path = find_object_md(args.object_dir)
        if md_path:
            yaml_data = load_yaml_from_md(md_path)
            print(f"  MD: {md_path.name}")
        else:
            print(f"  MD: не знайдено в Объекты/{args.object_dir}/wiki/objects/")

    # ── Формуємо subject dict ──
    subject: dict = {
        "name":              (args.name if args.name != "Об'єкт оцінки"
                              else yaml_data.get("name", args.name)),
        "type":              args.type,
        "area":              args.area or yaml_data.get("Area"),
        "zone":              yaml_data.get("Location_Zone") or args.zone,
        "district":          yaml_data.get("District"),
        "metro_min":         yaml_data.get("Distance_to_Metro"),
        "building_class":    yaml_data.get("Building_Class"),
        "floor_type":        yaml_data.get("Floor_Type"),
        "condition":         yaml_data.get("Condition_Type"),
        "renovation_style":  yaml_data.get("Renovation_Style"),
        "ceiling_height":    yaml_data.get("Ceiling_Height"),
        "generator":         yaml_data.get("Generator"),
        "shelter":           yaml_data.get("Shelter"),
        "parking_underground": yaml_data.get("Parking_Underground", 0),
        "parking_open":      yaml_data.get("Parking_Open", 0),
        "opex_owner":        yaml_data.get("OPEX_Owner"),
        "vacancy":           args.vacancy,
        "opex_pct":          args.opex_pct or OPEX_PCT.get(args.type, 0.20),
    }

    if not subject["area"]:
        sys.exit("Помилка: --area не задана і в MD не знайдено.")

    # ── Full Data Audit ──
    if not args.no_audit:
        subject = full_data_audit(subject, args.object_dir, md_path)

    # ── Динамічні вхідні ──
    if not args.no_interactive:
        discount, cap = ask_dynamic_inputs(args.type)
    else:
        discount = args.discount or DISCOUNT_RATES.get(args.type, 0.07)
        cap      = args.cap_rate or CAP_RATES.get(args.type, 0.10)
        print(f"  Торг: {discount*100:.1f}%  |  Cap Rate: {cap*100:.1f}%")

    # ── Аналоги — ВСІ ──
    analogs = load_all_analogs(args.type)
    print(f"  Аналогів завантажено: {len(analogs)}")
    if not analogs:
        print("  ⚠ Аналогів не знайдено. Перевірте Clippings/<Category>/")
        sys.exit(1)

    # ── Розрахунки ──
    print(f"\n[B1] Порівняльний аналіз орендних ставок ({len(analogs)} аналогів)...")
    b1 = rent_comparable_approach(subject, analogs, discount)
    if "error" in b1:
        sys.exit(f"B1: {b1['error']}")
    print(f"     Ринкова ставка (TRIMMEAN): ${b1['market_rent']:.2f}/м²/міс")
    print(f"     Діапазон: ${min(b1['finals']):.2f} – ${max(b1['finals']):.2f}/м²")

    print("[B2] Дохідний підхід...")
    b2 = income_approach(float(subject["area"]), b1["market_rent"],
                         subject["vacancy"], subject["opex_pct"], cap)
    print(f"     PGI: ${b2['pgi']:,.0f}  NOI: ${b2['noi']:,.0f}  "
          f"Cap {cap*100:.0f}%  →  ${b2['value']:,.0f}")

    b3 = reconciliation(b1["market_rent"] * float(subject["area"]),
                        b2["value"], args.weight_b1, args.weight_b2)

    # ── Генерація Excel ──
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    name_for_file = re.sub(r"[^\w\-_а-яА-ЯіІїЇєЄ ]", "_",
                           subject["name"]).strip().replace(" ", "_")[:40]
    xlsx_path = out_dir / f"Valuation_{name_for_file}.xlsx"

    print(f"\nГенерую Excel...")
    generate_excel(subject, b1, b2, b3, xlsx_path, discount, cap)
    print(f"  Збережено: {xlsx_path}")
    print(f"\n  Підсумок: ринкова ставка ${b1['market_rent']:.2f}/м²/міс  "
          f"|  NOI/Cap = ${b2['value']:,.0f}")


if __name__ == "__main__":
    main()
