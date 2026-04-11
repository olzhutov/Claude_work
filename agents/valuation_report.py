#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор звіту про оцінку комерційної нерухомості.
Методологія: .agents/skills/cre-valuation/ [B1] [B2] [B3]

Підходи:
  [B1] Порівняльний — матриця корегувань для 5+ аналогів
  [B2] Дохідний    — PGI → EGI → NOI → Cap Rate
  [B3] Узгодження  — зважена середня (за замовч. 50/50)

Запуск:
    python3 agents/valuation_report.py --name "Фастов завод" \
        --area 5000 --price 1500000 --type Warehouse \
        --rent-rate 5.5 --vacancy 0.15 --cap-rate 0.12

    python3 agents/valuation_report.py --help
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import date
from pathlib import Path
from typing import Any

import yaml

# ─── Залежності ──────────────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Встановіть: pip3 install openpyxl")

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    sys.exit("Встановіть: pip3 install python-docx")

# ─── Шляхи ───────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent.parent
CLIPPINGS_DIR = BASE_DIR / "Clippings"
OUTPUT_DIR    = BASE_DIR / "output" / "reports"

# ─── Ринкові константи ────────────────────────────────────────────────────────

# Знижка на торг [B1]: % від ціни пропозиції
DISCOUNT_RATES: dict[str, float] = {
    "Warehouse":   0.07,
    "Office":      0.08,
    "Retail":      0.06,
    "Land":        0.05,
    "Residential": 0.05,
}

# Ринкові Cap Rate [B2]: орієнтири Україна 2024-25
CAP_RATES: dict[str, float] = {
    "Warehouse":   0.12,
    "Office":      0.10,
    "Retail":      0.11,
    "Land":        0.08,
    "Residential": 0.08,
}

# OPEX як % від EGI [B2]
OPEX_PCT: dict[str, float] = {
    "Warehouse":   0.15,
    "Office":      0.20,
    "Retail":      0.18,
    "Land":        0.05,
    "Residential": 0.25,
}

# Числовий рейтинг зон для корегувань [B1]
ZONE_RANK: dict[str, int] = {
    "Center": 4, "Middle": 3, "Periphery": 2, "Suburbs": 1, "Unknown": 2,
}

# ─── Шаблонні аналоги (запасні якщо нема кліпінгів) ────────────────────────

DEMO_ANALOGS: dict[str, list[dict]] = {
    "Warehouse": [
        {"name": "Склад, Дарниця", "area": 5000, "price": 1350000,
         "price_psm": 270, "price_psm_adj": 251, "zone": "Periphery",
         "condition": "з ремонтом", "ceiling": 4.0, "power_kw": 160, "railway": True},
        {"name": "Вир.-склад. пр., Радистів", "area": 1500, "price": 480000,
         "price_psm": 320, "price_psm_adj": 298, "zone": "Periphery",
         "condition": "з ремонтом", "ceiling": 6.0, "power_kw": 400, "railway": False},
        {"name": "Склад. комплекс, Мила", "area": 10323, "price": 2500000,
         "price_psm": 242, "price_psm_adj": 225, "zone": "Suburbs",
         "condition": "з ремонтом", "ceiling": 7.0, "power_kw": 500, "railway": False},
        {"name": "Склад, Видубичі", "area": 3000, "price": 2800000,
         "price_psm": 933, "price_psm_adj": 858, "zone": "Center",
         "condition": "з ремонтом", "ceiling": 5.0, "power_kw": 300, "railway": False},
        {"name": "Логіст. центр, Бровари", "area": 8000, "price": 1200000,
         "price_psm": 150, "price_psm_adj": 140, "zone": "Suburbs",
         "condition": "після будівельників", "ceiling": 10.0, "power_kw": 800, "railway": True},
    ],
    "Office": [
        {"name": "Офіс, Хмельницького, 750м²", "area": 750, "price": 3750000,
         "price_psm": 5000, "price_psm_adj": 4600, "zone": "Center",
         "condition": "з ремонтом", "class": "A", "metro_min": 1},
        {"name": "БЦ, Голосіїв, 1200м²", "area": 1200, "price": 3600000,
         "price_psm": 3000, "price_psm_adj": 2760, "zone": "Middle",
         "condition": "з ремонтом", "class": "B+", "metro_min": 7},
        {"name": "Офіс, Поділ, 500м²", "area": 500, "price": 3000000,
         "price_psm": 6000, "price_psm_adj": 5520, "zone": "Center",
         "condition": "з ремонтом", "class": "B+", "metro_min": 5},
        {"name": "БЦ, Лук'янівка, 2000м²", "area": 2000, "price": 5000000,
         "price_psm": 2500, "price_psm_adj": 2300, "zone": "Middle",
         "condition": "з ремонтом", "class": "B", "metro_min": 10},
        {"name": "Офіс, Оболонь, 900м²", "area": 900, "price": 1980000,
         "price_psm": 2200, "price_psm_adj": 2024, "zone": "Middle",
         "condition": "під оздоблення", "class": "B", "metro_min": 8},
    ],
    "Retail": [
        {"name": "Магазин, Хрещатик, 300м²", "area": 300, "price": 3900000,
         "price_psm": 13000, "price_psm_adj": 12220, "zone": "Center",
         "condition": "з ремонтом"},
        {"name": "Рітейл, Оболонь, 450м²", "area": 450, "price": 2250000,
         "price_psm": 5000, "price_psm_adj": 4700, "zone": "Middle",
         "condition": "з ремонтом"},
        {"name": "Магазин, Троєщина, 600м²", "area": 600, "price": 1800000,
         "price_psm": 3000, "price_psm_adj": 2820, "zone": "Periphery",
         "condition": "з ремонтом"},
        {"name": "Стріт-рітейл, Поділ, 200м²", "area": 200, "price": 2600000,
         "price_psm": 13000, "price_psm_adj": 12220, "zone": "Center",
         "condition": "з ремонтом"},
        {"name": "Рітейл, Бровари, 800м²", "area": 800, "price": 1600000,
         "price_psm": 2000, "price_psm_adj": 1880, "zone": "Suburbs",
         "condition": "під оздоблення"},
    ],
}


# ─────────────────────────────────────────────────────────────────────────────
# Завантаження аналогів з Clippings/*.md
# ─────────────────────────────────────────────────────────────────────────────

FRONTMATTER_RE = re.compile(r"^---\n(.*?)\n---", re.DOTALL)


def _load_clipping_fm(path: Path) -> dict:
    """Читає YAML-frontmatter з MD-файлу кліпінгу."""
    try:
        text = path.read_text(encoding="utf-8")
        m = FRONTMATTER_RE.match(text)
        if m:
            return yaml.safe_load(m.group(1)) or {}
    except Exception:
        pass
    return {}


def load_analogs_from_clippings(obj_type: str, deal_type: str) -> list[dict]:
    """
    Шукає розпарсені кліпінги відповідної категорії та типу угоди.
    Фільтрує за полем Category у frontmatter (незалежно від назви папки).
    Повертає список dict з полями для матриці коригувань.
    """
    results = []
    pattern = "**/*.md"
    for md_path in sorted(CLIPPINGS_DIR.glob(pattern)):
        fm = _load_clipping_fm(md_path)
        if not fm.get("parsed"):
            continue
        if fm.get("Category") != obj_type:
            continue
        if fm.get("Deal_Type") != deal_type:
            continue

        analog: dict[str, Any] = {
            "name": md_path.stem[:50],
            "area": fm.get("Area"),
            "zone": fm.get("Location_Zone", "Periphery"),
            "condition": fm.get("Condition_Type", "з ремонтом"),
        }

        if deal_type == "Sale":
            analog["price"]       = fm.get("Price")
            analog["price_psm"]   = fm.get("Price_per_sqm")
            analog["price_psm_adj"] = fm.get("Price_per_sqm_Adjusted", fm.get("Price_per_sqm"))
        else:
            analog["price"]       = fm.get("Rent_Monthly_Total")
            analog["price_psm"]   = fm.get("Rent_per_sqm")
            analog["price_psm_adj"] = fm.get("Rent_Adjusted", fm.get("Rent_per_sqm"))

        # Тип-специфічні поля
        analog["ceiling"]  = fm.get("Ceiling_Height")
        analog["power_kw"] = fm.get("Power_kW")
        analog["railway"]  = fm.get("Railway", False)
        analog["class"]    = fm.get("Building_Class")
        analog["metro_min"] = fm.get("Distance_to_Metro")

        if analog.get("price_psm"):
            results.append(analog)

    return results


# ─────────────────────────────────────────────────────────────────────────────
# [B1] Порівняльний підхід — матриця корегувань
# ─────────────────────────────────────────────────────────────────────────────

def _area_adj(subject_area: float, analog_area: float | None) -> float:
    """
    Корегування на масштаб площі [B1].
    Більший аналог → він дешевший на одиницю (inferior for buyer) → коригуємо вгору.
    Крок: ±3% на кожні ±25% різниці площі.
    """
    if not analog_area or analog_area <= 0:
        return 0.0
    ratio = subject_area / analog_area
    if ratio > 1.5:   return +0.06
    if ratio > 1.25:  return +0.03
    if ratio < 0.67:  return -0.06
    if ratio < 0.80:  return -0.03
    return 0.0


def _zone_adj(subject_zone: str, analog_zone: str) -> float:
    """
    Корегування на локацію (зону) [B1].
    Аналог у кращій зоні → дорожчий → коригуємо вниз (-).
    """
    diff = ZONE_RANK.get(subject_zone, 2) - ZONE_RANK.get(analog_zone, 2)
    return diff * 0.08   # ±8% за рівень


def _condition_adj(analog_condition: str | None) -> float:
    """Корегування на стан (суб'єкт — з ремонтом як база)."""
    c = (analog_condition or "").lower()
    if "після будівельників" in c or "без ремонт" in c:
        return +0.15
    if "під оздоблення" in c:
        return +0.10
    if "з меблями" in c or "дизайнер" in c:
        return -0.05
    return 0.0


def _tech_adj_warehouse(subject: dict, analog: dict) -> float:
    """Технічні корегування для складу [B1]: висота, потужність, з/д."""
    adj = 0.0
    s_ceil = subject.get("ceiling_height", 6.0)
    a_ceil = analog.get("ceiling")
    if a_ceil:
        if a_ceil >= 10 and s_ceil < 10:  adj -= 0.08
        elif a_ceil >= 8  and s_ceil < 8: adj -= 0.05
        elif a_ceil < 5   and s_ceil >= 5: adj += 0.05
    if subject.get("railway") and not analog.get("railway"):
        adj += 0.05
    if not subject.get("railway") and analog.get("railway"):
        adj -= 0.05
    return adj


def _tech_adj_office(subject: dict, analog: dict) -> float:
    """Технічні корегування для офісу [B1]: клас будівлі, метро."""
    class_rank = {"A": 4, "B+": 3, "B": 2, "C": 1}
    s_class = class_rank.get(subject.get("building_class", "B"), 2)
    a_class = class_rank.get(analog.get("class"), 2)
    adj = (s_class - a_class) * 0.05

    s_metro = subject.get("metro_min", 10)
    a_metro = analog.get("metro_min") or 15
    if s_metro <= 3 and a_metro > 10: adj += 0.05
    if s_metro > 10 and a_metro <= 3: adj -= 0.05
    return adj


def comparative_approach(subject: dict, analogs: list[dict]) -> dict:
    """
    [B1] Порівняльний підхід — матриця корегувань.

    Args:
        subject: dict із параметрами об'єкта оцінки
        analogs: список аналогів (≥ 5)

    Returns:
        dict з таблицею, скоригованими цінами і підсумковою вартістю
    """
    rows = []
    obj_type = subject.get("type", "Warehouse")

    for a in analogs:
        psm_adj = a.get("price_psm_adj") or a.get("price_psm") or 0
        if not psm_adj:
            continue

        adj_scale    = _area_adj(subject["area"], a.get("area"))
        adj_location = _zone_adj(subject.get("zone", "Periphery"), a.get("zone", "Periphery"))
        adj_condition = _condition_adj(a.get("condition"))

        if obj_type == "Warehouse":
            adj_tech = _tech_adj_warehouse(subject, a)
        elif obj_type == "Office":
            adj_tech = _tech_adj_office(subject, a)
        else:
            adj_tech = 0.0

        total_adj = adj_scale + adj_location + adj_condition + adj_tech
        final_psm = psm_adj * (1 + total_adj)

        rows.append({
            "name":          a.get("name", "—"),
            "area":          a.get("area"),
            "price_psm":     a.get("price_psm"),
            "discount_adj":  a.get("price_psm_adj"),
            "adj_scale":     adj_scale,
            "adj_location":  adj_location,
            "adj_condition": adj_condition,
            "adj_tech":      adj_tech,
            "total_adj":     total_adj,
            "final_psm":     final_psm,
        })

    if not rows:
        return {"error": "Немає аналогів для порівняльного підходу"}

    final_psms = [r["final_psm"] for r in rows]
    avg_psm    = sum(final_psms) / len(final_psms)
    value_b1   = avg_psm * subject["area"]

    return {
        "rows":     rows,
        "avg_psm":  avg_psm,
        "value":    value_b1,
        "n":        len(rows),
    }


# ─────────────────────────────────────────────────────────────────────────────
# [B2] Дохідний підхід
# ─────────────────────────────────────────────────────────────────────────────

def income_approach(subject: dict) -> dict:
    """
    [B2] Дохідний підхід: PGI → EGI → NOI → Капіталізація.

    Args:
        subject: dict з area, rent_rate, vacancy_rate, cap_rate, type

    Returns:
        dict з проміжними показниками і вартістю
    """
    area         = subject["area"]
    rent_rate    = subject.get("rent_rate", 6.0)    # $/м²/міс
    vacancy      = subject.get("vacancy", 0.15)
    obj_type     = subject.get("type", "Warehouse")
    cap_rate     = subject.get("cap_rate") or CAP_RATES.get(obj_type, 0.11)
    opex_pct     = subject.get("opex_pct") or OPEX_PCT.get(obj_type, 0.18)

    pgi          = area * rent_rate * 12             # Потенційний валовий дохід
    vacancy_loss = pgi * vacancy
    egi          = pgi - vacancy_loss                # Ефективний валовий дохід
    opex         = egi * opex_pct
    noi          = egi - opex                        # Чистий операційний дохід
    value_b2     = noi / cap_rate

    return {
        "area":         area,
        "rent_rate":    rent_rate,
        "vacancy":      vacancy,
        "cap_rate":     cap_rate,
        "opex_pct":     opex_pct,
        "pgi":          pgi,
        "vacancy_loss": vacancy_loss,
        "egi":          egi,
        "opex":         opex,
        "noi":          noi,
        "value":        value_b2,
    }


# ─────────────────────────────────────────────────────────────────────────────
# [B3] Узгодження результатів
# ─────────────────────────────────────────────────────────────────────────────

def reconciliation(b1: dict, b2: dict,
                   weight_b1: float = 0.50,
                   weight_b2: float = 0.50) -> dict:
    """
    [B3] Узгодження вартості двох підходів.

    Args:
        b1, b2:   результати відповідних підходів
        weight_b1, weight_b2: ваги (в сумі = 1.0)

    Returns:
        dict з узгодженою вартістю та округленням
    """
    v1 = b1.get("value", 0)
    v2 = b2.get("value", 0)

    if v1 and not v2:
        return {"value": v1, "weight_b1": 1.0, "weight_b2": 0.0,
                "value_b1": v1, "value_b2": 0, "note": "Тільки B1"}
    if v2 and not v1:
        return {"value": v2, "weight_b1": 0.0, "weight_b2": 1.0,
                "value_b1": 0, "value_b2": v2, "note": "Тільки B2"}

    weighted = v1 * weight_b1 + v2 * weight_b2

    # Округлення до значущих цифр
    magnitude = 10 ** (len(str(int(weighted))) - 2)
    rounded = round(weighted / magnitude) * magnitude

    return {
        "value_b1":  v1,
        "value_b2":  v2,
        "weight_b1": weight_b1,
        "weight_b2": weight_b2,
        "weighted":  weighted,
        "value":     rounded,
        "note":      f"Зважена ({weight_b1*100:.0f}% B1 + {weight_b2*100:.0f}% B2), "
                     f"округлено до {magnitude:,}",
    }


# ─────────────────────────────────────────────────────────────────────────────
# Стилі Excel
# ─────────────────────────────────────────────────────────────────────────────

_C_DARK  = "1C2E44"
_C_MID   = "2D4A6B"
_C_ACCENT = "D5B58A"
_C_LIGHT = "EAF0F6"
_C_WHITE = "FFFFFF"
_C_GREEN = "D5E8D4"
_C_YELL  = "FFF2CC"

_thin = Side(style="thin", color="AAAAAA")
_thick = Side(style="medium", color=_C_DARK)
_border_all = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
_border_hdr = Border(left=_thick, right=_thick, top=_thick, bottom=_thick)


def _hdr(ws, row: int, col: int, text: str, width: float | None = None,
          dark: bool = True, font_size: int = 10) -> None:
    """Встановлює стиль заголовочної комірки."""
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = Font(name="Calibri", bold=True, color=_C_WHITE if dark else _C_DARK,
                     size=font_size)
    cell.fill = PatternFill("solid", fgColor=_C_DARK if dark else _C_ACCENT)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = _border_all
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width


def _cell(ws, row: int, col: int, value: Any,
          fmt: str = "General", bold: bool = False, bg: str = _C_WHITE) -> None:
    """Записує звичайну комірку з форматуванням."""
    cell = ws.cell(row=row, column=col, value=value)
    cell.number_format = fmt
    cell.font = Font(name="Calibri", bold=bold, size=10)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _border_all


def _pct(v: float) -> str:
    return f"{v:+.1%}" if v != 0 else "0%"


# ─────────────────────────────────────────────────────────────────────────────
# Генерація Excel [B1 + B2 + B3]
# ─────────────────────────────────────────────────────────────────────────────

def generate_excel(subject: dict, b1: dict, b2: dict, b3: dict,
                   out_path: Path) -> None:
    """
    Створює Excel-файл з трьома аркушами:
      1. Об'єкт оцінки
      2. Порівняльний підхід (матриця корегувань)
      3. Дохідний підхід + Узгодження
    """
    wb = openpyxl.Workbook()

    # ── Аркуш 1: Об'єкт оцінки ───────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Об'єкт оцінки"
    ws1.sheet_view.showGridLines = False
    ws1.row_dimensions[1].height = 30

    _hdr(ws1, 1, 1, "ОБ'ЄКТ ОЦІНКИ", font_size=13)
    ws1.merge_cells("A1:C1")

    fields = [
        ("Назва об'єкта",    subject.get("name", "—")),
        ("Тип нерухомості",  subject.get("type", "—")),
        ("Адреса",           subject.get("address", "—")),
        ("Площа (GBA), м²",  subject["area"]),
        ("Зона розташування", subject.get("zone", "—")),
        ("Стан",             subject.get("condition", "з ремонтом")),
        ("Ціна пропозиції, $", subject.get("asking_price", "—")),
        ("Орендна ставка, $/м²/міс", subject.get("rent_rate", "—")),
        ("Вакансія, %",      f"{subject.get('vacancy', 0.15)*100:.0f}%"),
        ("Дата оцінки",      str(date.today())),
    ]

    for i, (label, val) in enumerate(fields, start=3):
        _cell(ws1, i, 1, label, bold=True, bg=_C_LIGHT)
        _cell(ws1, i, 2, val)
        ws1.cell(row=i, column=3)

    ws1.column_dimensions["A"].width = 30
    ws1.column_dimensions["B"].width = 25
    ws1.column_dimensions["C"].width = 10

    # ── Аркуш 2: Порівняльний підхід (B1) ────────────────────────────────────
    ws2 = wb.create_sheet("B1 Порівняльний підхід")
    ws2.sheet_view.showGridLines = False

    headers_b1 = [
        ("Аналог",              22),
        ("Площа, м²",           10),
        ("Ціна пропозиції,\n$/м²", 12),
        ("Після торгу,\n$/м²",  12),
        ("Кор. масштаб",        10),
        ("Кор. локація",        10),
        ("Кор. стан",           10),
        ("Кор. техн.",          10),
        ("Σ корегувань",        11),
        ("Скоригована\nціна, $/м²", 14),
    ]

    ws2.row_dimensions[1].height = 10
    ws2.row_dimensions[2].height = 40
    for col, (label, width) in enumerate(headers_b1, start=1):
        _hdr(ws2, 2, col, label, width=width)

    # Заголовок аркуша
    ws2.cell(row=1, column=1, value="[B1] ПОРІВНЯЛЬНИЙ ПІДХІД — МАТРИЦЯ КОРЕГУВАНЬ")
    ws2.cell(row=1, column=1).font = Font(bold=True, size=11, color=_C_DARK)
    ws2.merge_cells(f"A1:{get_column_letter(len(headers_b1))}1")

    rows_b1 = b1.get("rows", [])
    for i, row in enumerate(rows_b1, start=3):
        bg = _C_WHITE if i % 2 == 1 else _C_LIGHT
        _cell(ws2, i, 1, row["name"],          bold=False, bg=bg)
        _cell(ws2, i, 2, row.get("area"),       fmt="#,##0", bg=bg)
        _cell(ws2, i, 3, row.get("price_psm"),  fmt="#,##0", bg=bg)
        _cell(ws2, i, 4, row.get("discount_adj"), fmt="#,##0", bg=bg)
        _cell(ws2, i, 5, _pct(row["adj_scale"]),    bg=bg)
        _cell(ws2, i, 6, _pct(row["adj_location"]), bg=bg)
        _cell(ws2, i, 7, _pct(row["adj_condition"]), bg=bg)
        _cell(ws2, i, 8, _pct(row["adj_tech"]),     bg=bg)
        total = row["total_adj"]
        _cell(ws2, i, 9, _pct(total), bold=True,
              bg=_C_GREEN if abs(total) <= 0.15 else _C_YELL)
        _cell(ws2, i, 10, round(row["final_psm"], 0), fmt="#,##0 $", bold=True, bg=bg)

    # Підсумок B1
    sum_row = len(rows_b1) + 4
    ws2.row_dimensions[sum_row].height = 22
    _cell(ws2, sum_row, 1, "СЕРЕДНЯ скоригована ціна, $/м²",
          bold=True, bg=_C_ACCENT)
    ws2.merge_cells(f"A{sum_row}:I{sum_row}")
    _cell(ws2, sum_row, 10, round(b1.get("avg_psm", 0), 0),
          fmt="#,##0 $", bold=True, bg=_C_ACCENT)

    sum_row2 = sum_row + 1
    _cell(ws2, sum_row2, 1, f"ВАРТІСТЬ за порівняльним підходом (× {subject['area']:,.0f} м²)",
          bold=True, bg=_C_DARK)
    ws2.cell(row=sum_row2, column=1).font = Font(bold=True, color=_C_WHITE, size=10)
    ws2.merge_cells(f"A{sum_row2}:I{sum_row2}")
    _cell(ws2, sum_row2, 10, round(b1.get("value", 0), 0),
          fmt="#,##0 $", bold=True, bg=_C_DARK)
    ws2.cell(row=sum_row2, column=10).font = Font(bold=True, color=_C_WHITE, size=10)

    # ── Аркуш 3: Дохідний підхід + Узгодження ────────────────────────────────
    ws3 = wb.create_sheet("B2+B3 Дохідний та Узгодження")
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions["A"].width = 38
    ws3.column_dimensions["B"].width = 20
    ws3.column_dimensions["C"].width = 14

    def _section(r: int, title: str) -> int:
        ws3.row_dimensions[r].height = 24
        cell = ws3.cell(row=r, column=1, value=title)
        cell.font = Font(bold=True, size=11, color=_C_WHITE)
        cell.fill = PatternFill("solid", fgColor=_C_MID)
        cell.alignment = Alignment(vertical="center")
        ws3.merge_cells(f"A{r}:C{r}")
        return r + 1

    def _row(r: int, label: str, value: Any, fmt: str = "General",
             bold: bool = False, bg: str = _C_WHITE) -> int:
        _cell(ws3, r, 1, label, bold=bold, bg=bg)
        ws3.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
        _cell(ws3, r, 2, value, fmt=fmt, bold=bold, bg=bg)
        ws3.cell(row=r, column=3)
        return r + 1

    r = 2
    r = _section(r, "[B2] ДОХІДНИЙ ПІДХІД")
    r = _row(r, "Площа (GLA), м²",             b2["area"],         "#,##0")
    r = _row(r, "Орендна ставка, $/м²/міс",    b2["rent_rate"],    "0.00")
    r = _row(r, "Вакансія, %",                  b2["vacancy"],      "0%")
    r = _row(r, "OPEX від EGI, %",              b2["opex_pct"],     "0%")
    r = _row(r, "Cap Rate (ринковий), %",        b2["cap_rate"],     "0.0%")
    r += 1
    r = _row(r, "PGI  — Потенційний валовий дохід, $/рік",
             round(b2["pgi"]),        "#,##0 $", bold=True)
    r = _row(r, "  − Втрати від вакансії, $/рік",
             round(b2["vacancy_loss"]), "#,##0 $")
    r = _row(r, "EGI  — Ефективний валовий дохід, $/рік",
             round(b2["egi"]),        "#,##0 $", bold=True, bg=_C_LIGHT)
    r = _row(r, "  − OPEX (операційні витрати), $/рік",
             round(b2["opex"]),       "#,##0 $")
    r = _row(r, "NOI  — Чистий операційний дохід, $/рік",
             round(b2["noi"]),        "#,##0 $", bold=True, bg=_C_LIGHT)
    r += 1
    r = _row(r, "ВАРТІСТЬ (B2) = NOI / Cap Rate",
             round(b2["value"]),      "#,##0 $", bold=True, bg=_C_ACCENT)

    r += 2
    r = _section(r, "[B3] УЗГОДЖЕННЯ РЕЗУЛЬТАТІВ")
    r = _row(r, "Вартість за порівняльним підходом (B1), $",
             round(b3["value_b1"]),   "#,##0 $")
    r = _row(r, "Вартість за дохідним підходом (B2), $",
             round(b3["value_b2"]),   "#,##0 $")
    r = _row(r, f"Вага B1 / B2",
             f"{b3['weight_b1']*100:.0f}% / {b3['weight_b2']*100:.0f}%")
    r = _row(r, "Зважена вартість (до округлення), $",
             round(b3.get("weighted", b3["value"])), "#,##0 $")
    r += 1
    r = _row(r, "ПІДСУМКОВА ВАРТІСТЬ, $",
             b3["value"],             "#,##0 $", bold=True, bg=_C_DARK)
    ws3.cell(row=r - 1, column=1).font = Font(bold=True, color=_C_WHITE, size=11)
    ws3.cell(row=r - 1, column=2).font = Font(bold=True, color=_C_WHITE, size=11)

    wb.save(out_path)
    print(f"  ✓ Excel: {out_path.name}")


# ─────────────────────────────────────────────────────────────────────────────
# Генерація Word
# ─────────────────────────────────────────────────────────────────────────────

def _set_cell_bg(cell, hex_color: str) -> None:
    """Встановлює фон комірки таблиці Word через XML."""
    from docx.oxml.ns import qn
    from docx.oxml import parse_xml
    shading = parse_xml(
        f'<w:shd {cell._tc.nsmap.get("w", "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"")}:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'w:val="clear" w:color="auto" w:fill="{hex_color}"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading)


def _docx_heading(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0x1C, 0x2E, 0x44)


def _docx_table_2col(doc: Document, rows: list[tuple[str, str]],
                     col_widths: tuple[float, float] = (9.0, 7.5)) -> None:
    """Додає двоколонкову таблицю (параметр | значення)."""
    tbl = doc.add_table(rows=len(rows) + 1, cols=2)
    tbl.style = "Table Grid"

    hdr_cells = tbl.rows[0].cells
    hdr_cells[0].text = "Параметр"
    hdr_cells[1].text = "Значення"
    for cell in hdr_cells:
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for i, (label, val) in enumerate(rows, start=1):
        tbl.cell(i, 0).text = label
        tbl.cell(i, 1).text = str(val)

    # Ширина стовпців
    from docx.oxml.ns import qn
    for row in tbl.rows:
        for j, w in enumerate(col_widths):
            cell = row.cells[j]
            cell.width = Cm(w)


def _docx_analogs_table(doc: Document, rows_b1: list[dict]) -> None:
    """Таблиця аналогів [B1] у Word."""
    cols = ["Аналог", "Площа,\nм²", "Ціна,\n$/м²", "Після\nторгу", "Σ кор-нь",
            "Підсумк.\nціна, $/м²"]
    tbl = doc.add_table(rows=len(rows_b1) + 1, cols=len(cols))
    tbl.style = "Table Grid"

    for j, h in enumerate(cols):
        cell = tbl.rows[0].cells[j]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True

    for i, row in enumerate(rows_b1, start=1):
        vals = [
            row["name"],
            f"{row.get('area') or '—':,}" if row.get("area") else "—",
            f"{row.get('price_psm') or '—':,.0f}" if row.get("price_psm") else "—",
            f"{row.get('discount_adj') or '—':,.0f}" if row.get("discount_adj") else "—",
            _pct(row["total_adj"]),
            f"{row['final_psm']:,.0f}",
        ]
        for j, v in enumerate(vals):
            tbl.cell(i, j).text = str(v)


def generate_word(subject: dict, b1: dict, b2: dict, b3: dict,
                  out_path: Path) -> None:
    """
    Генерує Word-звіт [B1, B2, B3] українською мовою.
    Структура: Резюме → Методологія → B1 → B2 → B3 → Ризики → Висновок
    """
    doc = Document()

    # ── Поля сторінки ──────────────────────────────────────────────────────
    section = doc.sections[0]
    section.page_width  = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = section.right_margin = Cm(2.5)
    section.top_margin  = section.bottom_margin = Cm(2.0)

    obj_name  = subject.get("name", "Об'єкт")
    obj_type  = subject.get("type", "Warehouse")
    obj_area  = subject["area"]
    obj_zone  = subject.get("zone", "—")
    today_str = date.today().strftime("%d.%m.%Y")

    # ══ Титульний блок ════════════════════════════════════════════════════════
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(f"ЗВІТ ПРО ОЦІНКУ РИНКОВОЇ ВАРТОСТІ\n{obj_name.upper()}")
    run.bold      = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0x1C, 0x2E, 0x44)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run(f"Станом на {today_str}  ·  Тип: {obj_type}  ·  Площа: {obj_area:,.0f} м²")

    doc.add_paragraph()

    # ══ 1. Резюме ═════════════════════════════════════════════════════════════
    _docx_heading(doc, "1. Резюме (Executive Summary)", level=1)

    doc.add_paragraph(
        f"Об'єктом оцінки є {obj_type.lower()} загальною площею {obj_area:,.0f} м², "
        f"розташований у зоні «{obj_zone}» (Київ/область). "
        f"Оцінка виконана за двома підходами: порівняльним [B1] та дохідним [B2], "
        f"результати узгоджено за методологією [B3]."
    )

    summary_rows = [
        ("Підсумкова ринкова вартість",
         f"$ {b3['value']:,.0f}"),
        ("  — за порівняльним підходом (B1)",
         f"$ {b3['value_b1']:,.0f}"),
        ("  — за дохідним підходом (B2)",
         f"$ {b3['value_b2']:,.0f}"),
        ("Вартість за 1 м² (узгоджена)",
         f"$ {b3['value'] / obj_area:,.0f}"),
        ("NOI (рік)",
         f"$ {b2['noi']:,.0f}"),
        ("Cap Rate (ринковий)",
         f"{b2['cap_rate']*100:.1f}%"),
        ("Дата оцінки", today_str),
    ]
    _docx_table_2col(doc, summary_rows)

    doc.add_paragraph()

    # ══ 2. Методологія ════════════════════════════════════════════════════════
    _docx_heading(doc, "2. Методологія оцінки", level=1)

    doc.add_paragraph(
        "Оцінка проведена відповідно до методологічного посібника "
        ".agents/skills/cre-valuation/ та включає:"
    )
    for bullet in [
        "[B1] Порівняльний підхід — зіставлення із ринковими аналогами із застосуванням "
        "матриці кількісних корегувань (торг, масштаб, локація, стан, техпараметри).",
        "[B2] Дохідний підхід — пряма капіталізація NOI за ринковою ставкою Cap Rate. "
        "Розраховано PGI → EGI (вакансія) → NOI (OPEX).",
        "[B3] Узгодження — зважена середня між підходами з урахуванням наявності "
        "достатньої бази аналогів та якості орендних даних.",
    ]:
        p = doc.add_paragraph(bullet, style="List Bullet")
        p.paragraph_format.left_indent = Cm(0.5)

    doc.add_paragraph()

    # ══ 3. Порівняльний підхід [B1] ══════════════════════════════════════════
    _docx_heading(doc, "3. Порівняльний підхід [B1]", level=1)

    _docx_heading(doc, "3.1 Таблиця аналогів", level=2)
    doc.add_paragraph(
        f"Для аналізу відібрано {b1['n']} ринкових аналогів. "
        "Ціни пропозицій скориговано на знижку від торгу відповідно до "
        f"ринкового стандарту для {obj_type} в Україні 2024–25."
    )
    _docx_analogs_table(doc, b1.get("rows", []))

    doc.add_paragraph()
    _docx_heading(doc, "3.2 Матриця корегувань", level=2)
    doc.add_paragraph(
        "Кожен аналог скориговано за чотирма групами факторів:"
    )
    adj_desc = [
        ("Масштаб площі", "±3–6% залежно від різниці площі із суб'єктом оцінки"),
        ("Локація (зона)", "±8% за рівень зони (Center > Middle > Periphery > Suburbs)"),
        ("Стан об'єкта", "+10–15% якщо аналог потребує ремонту; –5% якщо кращий стан"),
        ("Техн. параметри",
         "Склад: висота стелі, залізнична гілка. Офіс: клас будівлі, відстань до метро"),
    ]
    _docx_table_2col(doc, adj_desc, col_widths=(5.0, 11.5))

    doc.add_paragraph()
    avg_psm = b1.get("avg_psm", 0)
    val_b1  = b1.get("value", 0)
    p = doc.add_paragraph()
    p.add_run("Результат [B1]: ").bold = True
    p.add_run(
        f"середня скоригована ціна — $ {avg_psm:,.0f}/м²; "
        f"вартість об'єкта (× {obj_area:,.0f} м²) = "
    )
    p.add_run(f"$ {val_b1:,.0f}").bold = True

    doc.add_paragraph()

    # ══ 4. Дохідний підхід [B2] ══════════════════════════════════════════════
    _docx_heading(doc, "4. Дохідний підхід [B2]", level=1)

    doc.add_paragraph(
        "Метод прямої капіталізації. Ринкова ставка оренди підтверджена "
        "актуальними пропозиціями на OLX/DOM.RIA. NOI розраховано з урахуванням "
        f"вакансії {b2['vacancy']*100:.0f}% та операційних витрат {b2['opex_pct']*100:.0f}% від EGI."
    )

    income_rows = [
        ("GLA (площа, що здається в оренду), м²",
         f"{b2['area']:,.0f}"),
        ("Орендна ставка, $/м²/міс",
         f"{b2['rent_rate']:.2f}"),
        ("PGI (Потенційний валовий дохід), $/рік",
         f"{b2['pgi']:,.0f}"),
        ("Вакансія та втрати від несплати, $/рік",
         f"({b2['vacancy_loss']:,.0f})"),
        ("EGI (Ефективний валовий дохід), $/рік",
         f"{b2['egi']:,.0f}"),
        ("OPEX (операційні витрати), $/рік",
         f"({b2['opex']:,.0f})"),
        ("NOI (Чистий операційний дохід), $/рік",
         f"{b2['noi']:,.0f}"),
        ("Cap Rate (ринковий)", f"{b2['cap_rate']*100:.1f}%"),
        ("Вартість (B2) = NOI / Cap Rate",
         f"$ {b2['value']:,.0f}"),
    ]
    _docx_table_2col(doc, income_rows)

    doc.add_paragraph()

    # ══ 5. Узгодження [B3] ═══════════════════════════════════════════════════
    _docx_heading(doc, "5. Узгодження результатів [B3]", level=1)

    doc.add_paragraph(
        "Ваги підходів визначено виходячи з: достатності бази аналогів, "
        "якості та актуальності орендних даних, типу угоди."
    )

    rec_rows = [
        ("Вартість за B1 (порівняльний), $",   f"{b3['value_b1']:,.0f}"),
        ("Вага B1",                              f"{b3['weight_b1']*100:.0f}%"),
        ("Вартість за B2 (дохідний), $",        f"{b3['value_b2']:,.0f}"),
        ("Вага B2",                              f"{b3['weight_b2']*100:.0f}%"),
        ("Зважена (до округлення), $",
         f"{b3.get('weighted', b3['value']):,.0f}"),
        ("ПІДСУМКОВА ВАРТІСТЬ, $",              f"{b3['value']:,.0f}"),
    ]
    _docx_table_2col(doc, rec_rows)

    doc.add_paragraph()
    note = doc.add_paragraph()
    note.add_run("Примітка: ").bold = True
    note.add_run(b3.get("note", ""))

    doc.add_paragraph()

    # ══ 6. Ризики ═════════════════════════════════════════════════════════════
    _docx_heading(doc, "6. Ризики та обмеження", level=1)

    risks = [
        "Воєнний стан: ринок нерухомості функціонує в умовах обмеженого попиту "
        "та зниженої ліквідності — знижки реальних угод можуть бути вищими за типові.",
        "Обмежена база аналогів: через закритість ринку частина аналогів взята із "
        "відкритих пропозицій (OLX), що може відображати завищені ціни продавця.",
        "Курс валют: оцінка виконана у доларах США; зміна курсу UAH/USD вплине "
        "на гривневий еквівалент вартості.",
        "Юридичні ризики: наявність обтяжень, незавершених судових справ або "
        "питань прав власності може суттєво знизити ринкову вартість.",
    ]
    for risk in risks:
        p = doc.add_paragraph(risk, style="List Bullet")
        p.paragraph_format.left_indent = Cm(0.5)

    doc.add_paragraph()

    # ══ 7. Висновок ═══════════════════════════════════════════════════════════
    _docx_heading(doc, "7. Висновок", level=1)

    p = doc.add_paragraph()
    p.add_run(
        f"За результатами комплексної оцінки із застосуванням порівняльного та "
        f"дохідного підходів, ринкова вартість об'єкта нерухомості «{obj_name}» "
        f"(тип: {obj_type}, площа: {obj_area:,.0f} м², зона: {obj_zone}) "
        f"станом на {today_str} складає:"
    )

    final_p = doc.add_paragraph()
    final_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = final_p.add_run(f"$ {b3['value']:,.0f}  (USD)")
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0x1C, 0x2E, 0x44)

    doc.add_paragraph()
    disc = doc.add_paragraph()
    disc.add_run("Застереження: ").bold = True
    disc.add_run(
        "Цей звіт підготовлено виключно в інформаційних цілях і не є офіційним "
        "висновком суб'єкта оціночної діяльності у розумінні Закону України "
        "«Про оцінку майна». Для юридично значущих операцій необхідна сертифікована оцінка."
    )

    doc.save(out_path)
    print(f"  ✓ Word:  {out_path.name}")


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def build_subject(args: argparse.Namespace) -> dict:
    """Збирає dict суб'єкта оцінки з аргументів CLI."""
    return {
        "name":          args.name,
        "type":          args.type,
        "area":          args.area,
        "zone":          args.zone,
        "address":       args.address,
        "condition":     args.condition,
        "asking_price":  args.price,
        "rent_rate":     args.rent_rate,
        "vacancy":       args.vacancy,
        "cap_rate":      args.cap_rate,
        "opex_pct":      args.opex_pct,
        # warehouse extras
        "ceiling_height": args.ceiling,
        "railway":        args.railway,
        # office extras
        "building_class": args.building_class,
        "metro_min":      args.metro_min,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Генератор звіту про оцінку КН [B1, B2, B3]"
    )
    parser.add_argument("--name",    default="Об'єкт оцінки",  help="Назва об'єкта")
    parser.add_argument("--type",    default="Warehouse",
                        choices=["Warehouse", "Office", "Retail", "Land", "Residential"])
    parser.add_argument("--area",    type=float, required=True, help="Площа GBA, м²")
    parser.add_argument("--price",   type=float, default=None,  help="Ціна пропозиції, $")
    parser.add_argument("--zone",    default="Periphery",
                        choices=["Center", "Middle", "Periphery", "Suburbs"])
    parser.add_argument("--address", default="—")
    parser.add_argument("--condition", default="з ремонтом")

    # B2 параметри
    parser.add_argument("--rent-rate",  type=float, default=None,
                        help="Орендна ставка $/м²/міс")
    parser.add_argument("--vacancy",    type=float, default=0.15,
                        help="Вакансія 0.0-1.0 (default: 0.15)")
    parser.add_argument("--cap-rate",   type=float, default=None,
                        help="Cap Rate 0.0-1.0 (default: ринковий)")
    parser.add_argument("--opex-pct",   type=float, default=None,
                        help="OPEX вiдсоток вiд EGI (default: ринковий)")

    # B3 ваги
    parser.add_argument("--weight-b1",  type=float, default=0.50)
    parser.add_argument("--weight-b2",  type=float, default=0.50)

    # Тип-специфічні
    parser.add_argument("--ceiling",        type=float, default=None)
    parser.add_argument("--railway",        action="store_true")
    parser.add_argument("--building-class", default=None,
                        choices=["A", "B+", "B", "C"])
    parser.add_argument("--metro-min",      type=int,   default=None)

    # Вихід
    parser.add_argument("--out-dir", default=str(OUTPUT_DIR))

    args = parser.parse_args()

    if abs(args.weight_b1 + args.weight_b2 - 1.0) > 0.01:
        sys.exit("Помилка: weight-b1 + weight-b2 мають дорівнювати 1.0")

    subject = build_subject(args)

    # ── Аналоги ──────────────────────────────────────────────────────────────
    clipping_analogs = load_analogs_from_clippings(args.type, "Sale")
    demo             = DEMO_ANALOGS.get(args.type, DEMO_ANALOGS["Warehouse"])

    if len(clipping_analogs) >= 5:
        analogs = clipping_analogs
        print(f"  Аналоги: {len(analogs)} кліпінгів")
    else:
        combined = clipping_analogs + [
            a for a in demo if a not in clipping_analogs
        ]
        analogs = combined[:max(len(combined), 5)]
        print(f"  Аналоги: {len(clipping_analogs)} кліпінгів + {len(analogs)-len(clipping_analogs)} демо")

    # Якщо rent_rate не вказано — беремо ринковий орієнтир
    if args.rent_rate is None:
        defaults = {"Warehouse": 5.5, "Office": 14.0, "Retail": 18.0,
                    "Land": 3.0, "Residential": 10.0}
        subject["rent_rate"] = defaults.get(args.type, 6.0)
        print(f"  ⚠️  Припущення: rent_rate = {subject['rent_rate']} $/м²/міс (ринковий орієнтир)")

    # ── Розрахунки ───────────────────────────────────────────────────────────
    print("\n[B1] Порівняльний підхід...")
    b1 = comparative_approach(subject, analogs)
    if "error" in b1:
        sys.exit(f"B1 помилка: {b1['error']}")
    print(f"     Аналогів: {b1['n']} | Середня ціна: ${b1['avg_psm']:,.0f}/м² | "
          f"Вартість: ${b1['value']:,.0f}")

    print("[B2] Дохідний підхід...")
    b2 = income_approach(subject)
    print(f"     PGI: ${b2['pgi']:,.0f} | NOI: ${b2['noi']:,.0f} | "
          f"Cap {b2['cap_rate']*100:.0f}% | Вартість: ${b2['value']:,.0f}")

    print("[B3] Узгодження...")
    b3 = reconciliation(b1, b2, args.weight_b1, args.weight_b2)
    print(f"     Підсумкова вартість: ${b3['value']:,.0f}\n")

    # ── Генерація файлів ──────────────────────────────────────────────────────
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    safe_name = re.sub(r"[^\w\-_а-яА-ЯіІїЇєЄ ]", "_", args.name).strip().replace(" ", "_")

    xlsx_path = out_dir / f"Report_Calc_{safe_name}.xlsx"
    docx_path = out_dir / f"Звіт_про_оцінку_{safe_name}.docx"

    generate_excel(subject, b1, b2, b3, xlsx_path)
    generate_word(subject, b1, b2, b3, docx_path)

    print(f"\n  Папка: {out_dir}")


if __name__ == "__main__":
    main()
