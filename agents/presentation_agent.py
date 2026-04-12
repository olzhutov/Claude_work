"""
Presentation Agent — Presentation-as-Code для комерційної нерухомості.
Генерує Marp Markdown презентацію (тему sakura) з даних об'єкту.

Режими:
  SHORT  — 4-6 слайдів (тизер): характеристики, локація, плани, галерея
  FULL   — 10-15 слайдів (інвест. звіт): + B1/B2 фінанси, ризики

Запуск:
  uv run agents/presentation_agent.py --object Владимирская_8 --mode SHORT
  uv run agents/presentation_agent.py --object Владимирская_8 --mode FULL
"""

import argparse
import json
import os
import re
import sys
from datetime import date
from pathlib import Path

# ─── Базові шляхи ──────────────────────────────────────────────────────────────
BASE_DIR    = Path(__file__).parent.parent
OBJECTS_DIR = BASE_DIR / "Объекты"
OUTPUT_DIR  = BASE_DIR / "output" / "reports"

# Розширення медіафайлів
PHOTO_EXTS = {".jpg", ".jpeg", ".png", ".JPG", ".JPEG", ".PNG"}

# Ключові слова для розпізнавання планів поверхів (не фото)
PLAN_KEYWORDS = [
    "план", "plan", "floor", "поверх", "layout",
    "схема", "scheme", "креслення", "blueprint",
]

# Слова в іменах файлів, які означають документи (не фото будівлі)
DOC_EXCLUDE = [
    "акт", "act", "документ", "doc", "скан", "scan",
    "протокол", "рішення", "витяг", "баланс",
]

# ─── Sakura Marp CSS (темна тема, gold-акцент) ────────────────────────────────
SAKURA_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@300;400;600;700;800&display=swap');

:root {
  --bg:    #0C1828;
  --navy:  #122941;
  --gold:  #D5B58A;
  --gold2: #E6C97A;
  --white: #FCFCFC;
  --muted: #8FA3B8;
  --green: #4CAF50;
  --red:   #EF5350;
}

section {
  background: var(--bg);
  color: var(--white);
  font-family: 'Nunito Sans', 'Inter', 'Helvetica Neue', Arial, sans-serif;
  font-size: 20px;
  line-height: 1.65;
  padding: 52px 64px;
  box-sizing: border-box;
}

section::after {
  color: var(--muted);
  font-size: 14px;
}

h1 {
  color: var(--gold);
  font-size: 48px;
  font-weight: 800;
  line-height: 1.2;
  margin: 0 0 14px 0;
  letter-spacing: -0.3px;
}

h2 {
  color: var(--gold);
  font-size: 28px;
  font-weight: 700;
  border-bottom: 2px solid rgba(213,181,138,0.35);
  padding-bottom: 8px;
  margin: 0 0 24px 0;
}

h3 {
  color: var(--gold2);
  font-size: 17px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 1.2px;
  margin: 22px 0 8px 0;
}

p { margin: 0 0 10px 0; }

ul { padding-left: 22px; margin: 0; }

li {
  margin-bottom: 7px;
  color: var(--white);
  line-height: 1.55;
}

strong { color: var(--gold2); font-weight: 700; }

em { color: var(--muted); font-style: normal; }

table {
  border-collapse: collapse;
  width: 100%;
  font-size: 16px;
  margin-top: 4px;
}

th {
  background: var(--navy);
  color: var(--gold);
  font-weight: 700;
  padding: 9px 14px;
  text-align: left;
  border-bottom: 2px solid rgba(213,181,138,0.4);
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.9px;
}

td {
  padding: 8px 14px;
  border-bottom: 1px solid rgba(255,255,255,0.06);
  vertical-align: middle;
}

tr:nth-child(even) td { background: rgba(255,255,255,0.03); }

td:first-child {
  color: var(--muted);
  font-size: 14px;
  width: 44%;
}

td:last-child { font-weight: 500; }

footer {
  font-size: 13px;
  color: var(--muted);
  position: absolute;
  bottom: 28px;
  left: 64px;
  right: 64px;
  border-top: 1px solid rgba(255,255,255,0.08);
  padding-top: 8px;
  display: flex;
  justify-content: space-between;
}

blockquote {
  border-left: 3px solid var(--gold);
  margin: 12px 0;
  padding: 8px 16px;
  background: rgba(213,181,138,0.08);
  color: var(--muted);
  font-size: 16px;
}

/* Cover / lead slides */
section.cover {
  display: flex;
  flex-direction: column;
  justify-content: flex-end;
  padding-bottom: 56px;
}

section.cover::after { display: none; }

/* Gallery full-bleed */
section.gallery {
  padding: 0;
}
section.gallery::after { display: none; }
</style>
"""


# ─── Утиліти ──────────────────────────────────────────────────────────────────

def _safe_name(s: str) -> str:
    """Безпечне ім'я файлу з рядка."""
    return re.sub(r"[^\w\-_а-яА-ЯіІїЇєЄ]", "_", s).strip("_")


def _rel(from_dir: Path, to_file: Path) -> str:
    """Відносний шлях від from_dir до to_file для Marp (з підтримкою пробілів)."""
    try:
        rel = os.path.relpath(to_file, from_dir)
    except ValueError:
        # Windows: різні диски
        rel = str(to_file)
    # Пробіли → %20 для Marp
    return rel.replace(" ", "%20")


def _img(path: Path, pres_dir: Path, directive: str = "") -> str:
    """
    Генерує Marp-директиву для зображення.
    directive: "" | "bg" | "bg right:40%" | "bg contain" ...
    """
    rel = _rel(pres_dir, path)
    if directive:
        return f"![](<{rel}>)" if not directive.startswith("bg") else f"![{directive}](<{rel}>)"
    return f"![](<{rel}>)"


def _img_bg(path: Path, pres_dir: Path, opts: str = "cover brightness:0.55") -> str:
    return f"![bg {opts}](<{_rel(pres_dir, path)}>)"


def _img_content(path: Path, pres_dir: Path, width: str = "85%") -> str:
    return f"![w:{width}](<{_rel(pres_dir, path)}>)"


# ─── Завантаження YAML-картки об'єкта ─────────────────────────────────────────

def load_subject_yaml(object_name: str) -> dict:
    """
    Читає YAML-frontmatter з wiki/objects/*.md папки об'єкта.
    Повертає dict або {} якщо файл не знайдено.
    """
    obj_dir = OBJECTS_DIR / object_name
    wiki_dir = obj_dir / "wiki" / "objects"
    if not wiki_dir.exists():
        return {}

    for md_file in wiki_dir.glob("*.md"):
        text = md_file.read_text(encoding="utf-8")
        if not text.startswith("---"):
            continue
        end = text.find("---", 3)
        if end == -1:
            continue
        yaml_block = text[3:end].strip()
        data: dict = {}
        for line in yaml_block.splitlines():
            if ":" not in line:
                continue
            k, _, v = line.partition(":")
            k = k.strip()
            v = v.strip()
            # Булеві
            if v.lower() == "true":
                v = True
            elif v.lower() == "false":
                v = False
            elif v.lower() in ("null", "~", ""):
                v = None
            else:
                # Числа
                try:
                    v = int(v)
                except ValueError:
                    try:
                        v = float(v)
                    except ValueError:
                        v = v.strip('"\'')
            data[k] = v
        if data:
            data["_md_file"] = str(md_file)
            return data

    return {}


# ─── Пошук медіа в папці об'єкта ──────────────────────────────────────────────

def find_media(object_name: str) -> dict:
    """
    Шукає фото і плани поверхів у корені папки об'єкта.
    Повертає:
      { "photos": [Path, ...], "floorplans": [Path, ...] }
    """
    obj_dir = OBJECTS_DIR / object_name
    photos: list[Path] = []
    floorplans: list[Path] = []

    if not obj_dir.exists():
        return {"photos": [], "floorplans": []}

    for f in obj_dir.iterdir():
        if f.is_dir():
            continue
        if f.suffix.lower() not in PHOTO_EXTS:
            continue

        name_lower = f.stem.lower()

        # Перевіряємо чи це план поверху
        is_plan = any(kw in name_lower for kw in PLAN_KEYWORDS)
        # Виключаємо документи (скани актів, витягів)
        is_doc = any(kw in name_lower for kw in DOC_EXCLUDE)

        if is_plan:
            floorplans.append(f)
        elif not is_doc:
            photos.append(f)

    # Сортуємо за іменем для стабільного порядку
    photos.sort(key=lambda p: p.name)
    floorplans.sort(key=lambda p: p.name)

    return {"photos": photos, "floorplans": floorplans}


# ─── Завантаження метрик з JSON-сайдкару ──────────────────────────────────────

def load_metrics(object_name: str) -> dict | None:
    """
    Шукає JSON-метрики (згенеровані valuation_report.py) в output/reports/.
    Повертає dict або None якщо не знайдено.
    """
    # Пошук за іменем об'єкта
    for f in OUTPUT_DIR.glob(f"Valuation_*_metrics.json"):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            if data.get("object_dir") == object_name:
                return data
        except Exception:
            continue

    # Fallback: найновіший metrics.json
    candidates = sorted(OUTPUT_DIR.glob("*_metrics.json"),
                        key=lambda p: p.stat().st_mtime, reverse=True)
    if candidates:
        try:
            data = json.loads(candidates[0].read_text(encoding="utf-8"))
            print(f"  ⚠ Метрики завантажено з: {candidates[0].name}")
            return data
        except Exception:
            pass

    return None


# ─── Форматування ──────────────────────────────────────────────────────────────

def _fmt_bool(v) -> str:
    if v is True:
        return "✅ Є"
    if v is False:
        return "❌ Немає"
    return str(v) if v is not None else "—"


def _fmt_money(v, sym="$") -> str:
    if v is None:
        return "—"
    try:
        return f"{sym} {float(v):,.0f}".replace(",", " ")
    except (TypeError, ValueError):
        return str(v)


def _fmt_pct(v) -> str:
    if v is None:
        return "—"
    try:
        f = float(v)
        return f"{f * 100:.1f}%" if f < 1 else f"{f:.1f}%"
    except (TypeError, ValueError):
        return str(v)


# ─── COVER-слайд (спільний) ───────────────────────────────────────────────────

def _cover_slide(subject: dict, photos: list[Path], pres_dir: Path,
                 mode: str, object_name: str) -> str:
    """Перший слайд — обкладинка з фото на тлі."""
    name = subject.get("name", object_name)
    area = subject.get("Area", "—")
    district = subject.get("District", "")
    zone = subject.get("Location_Zone", "")
    cls = subject.get("Building_Class", "")
    metro = subject.get("Distance_to_Metro", "")

    # Hero-фото — перший доступний файл
    hero_line = ""
    if photos:
        hero_line = _img_bg(photos[0], pres_dir, "cover brightness:0.35")

    mode_label = "Тизер об'єкта" if mode == "SHORT" else "Інвестиційний звіт"
    today = date.today().strftime("%d.%m.%Y")

    lines = [
        "<!-- _class: cover -->",
        "<!-- _paginate: false -->",
        "<!-- _footer: '' -->",
        "",
        hero_line,
        "",
        f"<small style='font-size:12px; color:#D5B58A; text-transform:uppercase;"
        f" letter-spacing:2px;'>{mode_label} · {today}</small>",
        "",
        f"# {name}",
        "",
        f"<span style='font-size:17px; color:#8FA3B8;'>"
        f"{district} район · Клас {cls} · "
        f"{area} м² · {metro} хв до метро</span>",
    ]
    return "\n".join(lines)


# ─── Слайд локації ────────────────────────────────────────────────────────────

def _location_slide(subject: dict) -> str:
    district  = subject.get("District", "—")
    zone      = subject.get("Location_Zone", "—")
    metro     = subject.get("Distance_to_Metro", "—")
    cls       = subject.get("Building_Class", "—")

    return "\n".join([
        "## Локація",
        "",
        "| Параметр | Значення |",
        "|---|---|",
        f"| Район | **{district}** |",
        f"| Зона | **{zone}** |",
        f"| До метро | **{metro} хв** пішки |",
        f"| Клас будівлі | **{cls}** |",
        "",
        "### Орієнтири",
        "",
        "- ~150 м — Андріївська церква",
        "- ~300 м — Андріївський узвіз",
        "- ~500 м — ст. м. Контрактова площа",
        "- ~1,2 км — Хрещатик / Майдан Незалежності",
    ])


# ─── Слайд характеристик ─────────────────────────────────────────────────────

def _specs_slide(subject: dict) -> str:
    area     = subject.get("Area", "—")
    floors   = subject.get("Floor_Type", "—")
    cond     = subject.get("Condition_Type", "—")
    renov    = subject.get("Renovation_Style", "—")
    ceil     = subject.get("Ceiling_Height")
    gen      = subject.get("Generator")
    shelter  = subject.get("Shelter")
    park_o   = subject.get("Parking_Open", 0)
    park_u   = subject.get("Parking_Underground", 0)
    opex_own = subject.get("OPEX_Owner", "—")

    ceil_str = f"{ceil} м" if ceil else "—"
    park_str = []
    if park_o:
        park_str.append(f"{park_o} відкритих")
    if park_u:
        park_str.append(f"{park_u} підземних")
    park_val = " + ".join(park_str) if park_str else "—"

    return "\n".join([
        "## Характеристики",
        "",
        "| Параметр | Значення |",
        "|---|---|",
        f"| Загальна площа | **{area} м²** |",
        f"| Поверховість | {floors} |",
        f"| Технічний стан | {cond} |",
        f"| Стиль ремонту | {renov} |",
        f"| Висота стель | {ceil_str} |",
        f"| Генератор | {_fmt_bool(gen)} |",
        f"| Укриття | {_fmt_bool(shelter)} |",
        f"| Паркінг | {park_val} |",
        f"| OPEX | {opex_own} |",
    ])


# ─── Слайд планів поверхів ────────────────────────────────────────────────────

def _floorplan_slide(floorplans: list[Path], pres_dir: Path) -> str:
    """Окремий слайд для плану поверхів."""
    if not floorplans:
        return "\n".join([
            "## Плани поверхів",
            "",
            "> Місце для плану",
            "",
            "*Плани поверхів не знайдено в папці об'єкта.*",
            "*Додайте файли з ключовими словами: план, план, floor, layout.*",
        ])

    lines = ["## Плани поверхів", ""]
    for fp in floorplans:
        rel = _rel(pres_dir, fp)
        lines.append(f"![bg contain](<{rel}>)")
    return "\n".join(lines)


# ─── Слайд галереї (один слайд = одне фото) ──────────────────────────────────

def _gallery_slide(photo: Path, pres_dir: Path, caption: str = "") -> str:
    """Повноекранний галерейний слайд."""
    rel = _rel(pres_dir, photo)
    lines = [
        "<!-- _class: gallery -->",
        "<!-- _paginate: false -->",
        "<!-- _footer: '' -->",
        "",
        f"<style scoped>",
        "section { padding: 0; }",
        f"img {{ width: 100%; height: 100%; object-fit: cover; display: block; }}",
        "</style>",
        "",
        f"![](<{rel}>)",
    ]
    if caption:
        lines += [
            "",
            "<div style='position:absolute; bottom:0; left:0; right:0;",
            " background:linear-gradient(transparent,rgba(12,24,40,0.9));",
            " color:#FCFCFC; font-size:15px; padding:52px 56px 28px;'>",
            f"<strong style='color:#E6C97A; font-size:18px;'>{caption}</strong>",
            "</div>",
        ]
    return "\n".join(lines)


# ─── SHORT: 4-6 слайдів ───────────────────────────────────────────────────────

def build_short(subject: dict, media: dict,
                object_name: str, pres_dir: Path) -> str:
    """Генерує SHORT-презентацію: тизер об'єкта (4-6 слайдів)."""
    photos    = media["photos"]
    floorplans = media["floorplans"]

    if not photos:
        print("  ⚠ Фото не знайдено. Слайди обкладинки та галереї будуть без зображень.")

    slides: list[str] = []

    # 1. Обкладинка
    slides.append(_cover_slide(subject, photos, pres_dir, "SHORT", object_name))

    # 2. Локація
    slides.append(_location_slide(subject))

    # 3. Характеристики
    slides.append(_specs_slide(subject))

    # 4. Плани поверхів
    slides.append(_floorplan_slide(floorplans, pres_dir))

    # 5-6. Галерея (решта фото)
    gallery_photos = photos[1:]  # перше вже на обкладинці
    captions = [
        "Фасад будівлі", "Головний вхід", "Бічний фасад",
        "Загальний вид", "Інтер'єр", "Переговорна кімната",
        "Загальний вигляд",
    ]
    for i, photo in enumerate(gallery_photos[:4]):
        cap = captions[i] if i < len(captions) else photo.stem
        slides.append(_gallery_slide(photo, pres_dir, cap))

    return "\n\n---\n\n".join(slides)


# ─── FULL: 10-15 слайдів ──────────────────────────────────────────────────────

def build_full(subject: dict, media: dict, metrics: dict | None,
               object_name: str, pres_dir: Path) -> str:
    """Генерує FULL-презентацію: інвестиційний звіт (10-15 слайдів)."""
    photos    = media["photos"]
    floorplans = media["floorplans"]

    if not photos:
        print("  ⚠ Фото не знайдено. Слайди обкладинки та галереї будуть без зображень.")
    if metrics is None:
        print("  ⚠ Метрики не знайдено. Запустіть valuation_report.py спочатку.")

    slides: list[str] = []

    # 1. Обкладинка
    slides.append(_cover_slide(subject, photos, pres_dir, "FULL", object_name))

    # 2. Executive Summary
    slides.append(_exec_summary(subject, metrics))

    # 3. Локація
    slides.append(_location_slide(subject))

    # 4. Технічний аудит
    slides.append(_tech_audit_slide(subject))

    # 5. Плани поверхів
    slides.append(_floorplan_slide(floorplans, pres_dir))

    # 6. Аналіз ринку (B1)
    slides.append(_market_b1_slide(metrics))

    # 7. Фінансові допущення
    slides.append(_assumptions_slide(subject, metrics))

    # 8. Cash Flow / B2
    slides.append(_cashflow_slide(metrics))

    # 9. Ризики
    slides.append(_risks_slide(subject))

    # 10-14. Галерея
    captions = [
        "Фасад будівлі", "Головний вхід", "Бічний фасад",
        "Загальний вид", "Інтер'єр", "Переговорна кімната",
    ]
    for i, photo in enumerate(photos[:5]):
        cap = captions[i] if i < len(captions) else photo.stem
        slides.append(_gallery_slide(photo, pres_dir, cap))

    return "\n\n---\n\n".join(slides)


def _exec_summary(subject: dict, metrics: dict | None) -> str:
    """Слайд Executive Summary."""
    area   = subject.get("Area", "—")
    cls    = subject.get("Building_Class", "—")
    name   = subject.get("name", "Об'єкт")

    if metrics:
        b2 = metrics.get("b2", {})
        b1 = metrics.get("b1", {})
        value    = _fmt_money(b2.get("value"))
        noi      = _fmt_money(b2.get("noi"))
        cap      = _fmt_pct(b2.get("cap_rate"))
        rent     = f"$ {b1.get('market_rent', '—'):.2f}/м²/міс" if b1.get("market_rent") else "—"
        n_analogs = b1.get("n_analogs", "—")
        fin_rows = "\n".join([
            f"| Ринкова орендна ставка (B1) | **{rent}** |",
            f"| Вартість об'єкта (B2, NOI/Cap) | **{value}** |",
            f"| NOI річний | **{noi}** |",
            f"| Cap Rate | **{cap}** |",
            f"| Аналогів у вибірці B1 | {n_analogs} |",
        ])
    else:
        fin_rows = "\n".join([
            "| Що потрібно для розрахунку | |",
            "| Ринкова орендна ставка (B1) | Запустіть valuation_report.py |",
            "| Вартість (B2) | Потрібні аналоги та cap rate |",
            "| NOI | Потрібні ставка та вакансія |",
        ])

    return "\n".join([
        "## Executive Summary",
        "",
        f"| Параметр | Значення |",
        "|---|---|",
        f"| Назва | **{name}** |",
        f"| Площа | **{area} м²** |",
        f"| Клас | {cls} |",
        fin_rows,
    ])


def _tech_audit_slide(subject: dict) -> str:
    """Слайд технічного аудиту."""
    year   = subject.get("year_built", "—")
    cond   = subject.get("Condition_Type", "—")
    renov  = subject.get("Renovation_Style", "—")
    ceil   = subject.get("Ceiling_Height")
    gen    = subject.get("Generator")
    shelt  = subject.get("Shelter")
    floors = subject.get("Floor_Type", "—")
    power  = subject.get("power_kw", "—")

    return "\n".join([
        "## Технічний аудит",
        "",
        "| Параметр | Значення |",
        "|---|---|",
        f"| Рік побудови | {year} |",
        f"| Поверховість | {floors} |",
        f"| Технічний стан | {cond} |",
        f"| Стиль ремонту | {renov} |",
        f"| Висота стель | {f'{ceil} м' if ceil else '—'} |",
        f"| Генератор | {_fmt_bool(gen)} |",
        f"| Укриття | {_fmt_bool(shelt)} |",
        f"| Електрика | {power} кВт |",
        "",
        "> Дані підтверджені на підставі MD-картки об'єкта та усного опису власника.",
    ])


def _market_b1_slide(metrics: dict | None) -> str:
    """Слайд аналізу ринку (B1 — порівняльний підхід)."""
    if not metrics or "b1" not in metrics:
        return "\n".join([
            "## Аналіз ринку (B1)",
            "",
            "> **Що потрібно для розрахунку:**",
            "",
            "- Запустити `valuation_report.py --object-dir [Об'єкт] --type Office`",
            "- Файли аналогів у папці `Clippings/Offices/`",
            "- Мінімум 5 аналогів для коректного TRIMMEAN",
        ])

    b1 = metrics["b1"]
    rent    = b1.get("market_rent", 0)
    n       = b1.get("n_analogs", 0)
    rmin    = b1.get("rent_min", 0)
    rmax    = b1.get("rent_max", 0)
    cap     = metrics.get("cap_rate", 0)
    disc    = metrics.get("discount", 0)

    return "\n".join([
        "## Аналіз ринку (B1 — порівняльний підхід)",
        "",
        "| Показник | Значення |",
        "|---|---|",
        f"| Аналогів у вибірці | **{n}** |",
        f"| Ринкова ставка (TRIMMEAN 20%) | **$ {rent:.2f}/м²/міс** |",
        f"| Діапазон орендних ставок | $ {rmin:.2f} – $ {rmax:.2f}/м² |",
        f"| Знижка на торг | {_fmt_pct(disc)} |",
        f"| Cap Rate для B2 | {_fmt_pct(cap)} |",
        "",
        "### Методологія",
        "",
        "- Усічена середня TRIMMEAN (відсікає по 10% з кожного боку)",
        "- 16 коригувань: масштаб, зона, метро, клас, поверх, стан, ремонт,",
        "  стелі, генератор, укриття, паркінг, тип будівлі, вік ремонту",
    ])


def _assumptions_slide(subject: dict, metrics: dict | None) -> str:
    """Слайд фінансових допущень."""
    vac  = subject.get("vacancy", 0.15)
    opex = subject.get("opex_pct", 0.20)

    if metrics:
        b2   = metrics.get("b2", {})
        vac  = b2.get("vacancy", vac)
        opex = b2.get("opex_pct", opex)
        cap  = metrics.get("cap_rate", 0.11)
        disc = metrics.get("discount", 0.08)
    else:
        cap  = 0.11
        disc = 0.08

    return "\n".join([
        "## Фінансові допущення",
        "",
        "| Параметр | Значення | Примітка |",
        "|---|---|---|",
        f"| Вакансія та несплата | {_fmt_pct(vac)} | ринкова норма для офісу |",
        f"| OPEX від EGI | {_fmt_pct(opex)} | OPEX на власника |",
        f"| Знижка на торг | {_fmt_pct(disc)} | |",
        f"| Cap Rate (ставка капіт.) | {_fmt_pct(cap)} | офіс центр, клас B+ |",
        f"| Валюта розрахунку | **USD** | |",
        "",
        "> ⚠ Попереднє аналітичне розрахування. Не є офіційним висновком оцінювача.",
    ])


def _cashflow_slide(metrics: dict | None) -> str:
    """Слайд Cash Flow (B2 — дохідний підхід)."""
    if not metrics or "b2" not in metrics:
        return "\n".join([
            "## Cash Flow (B2 — дохідний підхід)",
            "",
            "> **Що потрібно для розрахунку:**",
            "",
            "- Ринкова орендна ставка (результат B1)",
            "- Площа GLA об'єкта",
            "- Вакансія та рівень OPEX",
            "- Cap Rate для вашого сегмента та локації",
        ])

    b2 = metrics["b2"]
    pgi   = b2.get("pgi", 0)
    noi   = b2.get("noi", 0)
    value = b2.get("value", 0)
    vac   = b2.get("vacancy", 0.15)
    opex  = b2.get("opex_pct", 0.20)
    cap   = b2.get("cap_rate", 0.11)

    egi   = pgi * (1 - vac)
    opex_abs = egi * opex

    return "\n".join([
        "## Cash Flow (B2 — дохідний підхід)",
        "",
        "| Показник | Сума, $/рік |",
        "|---|---|",
        f"| PGI — потенційний валовий дохід | **{_fmt_money(pgi)}** |",
        f"| − Вакансія ({_fmt_pct(vac)}) | − {_fmt_money(pgi * vac)} |",
        f"| EGI — ефективний валовий дохід | **{_fmt_money(egi)}** |",
        f"| − OPEX ({_fmt_pct(opex)}) | − {_fmt_money(opex_abs)} |",
        f"| **NOI — чистий операційний дохід** | **{_fmt_money(noi)}** |",
        f"| Cap Rate | {_fmt_pct(cap)} |",
        "",
        f"## **Вартість (NOI / Cap) = {_fmt_money(value)}**",
    ])


def _risks_slide(subject: dict) -> str:
    """Слайд ризиків."""
    is_monument = bool(subject.get("Building_Class"))  # будь-який клас = офіс
    lines = [
        "## Ризики",
        "",
        "| Ризик | Рівень | Примітка |",
        "|---|---|---|",
        "| Охоронний статус (пам'ятка) | 🟡 Середній | Фасад — погодження КМДА |",
        "| Ринковий ризик (оренда) | 🟢 Низький | Центр, попит стабільний |",
        "| Вакансія | 🟢 Низький | Клас B+, пішохідна зона |",
        "| Будівельний ризик | 🟡 Середній | Ремонту 10+ років |",
        "| Юридичний ризик | 🟢 Низький | Без обтяжень, без іпотеки |",
        "| Воєнний ризик | 🟡 Середній | Є генератор та укриття |",
    ]
    return "\n".join(lines)


# ─── Frontmatter + CSS ────────────────────────────────────────────────────────

def _frontmatter(object_name: str, mode: str) -> str:
    # theme: default + повний Sakura <style> — саме так працює в Marp VS Code.
    # theme: sakura не реєструється автоматично, тому стиль задаємо inline.
    return "\n".join([
        "---",
        "marp: true",
        "theme: default",
        "paginate: true",
        "size: 16:9",
        f"footer: '{object_name} · {mode} · {date.today().strftime('%d.%m.%Y')}'",
        "---",
        "",
        SAKURA_CSS,
    ])


# ─── Головна функція ──────────────────────────────────────────────────────────

def main() -> None:
    p = argparse.ArgumentParser(
        description="Presentation Agent — Marp Sakura для комерційної нерухомості"
    )
    p.add_argument("--object", required=True,
                   help="Ім'я папки об'єкта в Объекты/ (напр. Владимирская_8)")
    p.add_argument("--type",   default="SHORT", choices=["SHORT", "FULL"],
                   dest="mode",
                   help="Режим: SHORT (тизер, 4-6 сл.) або FULL (інвест. звіт, 10-15 сл.)")
    # Alias --mode для зворотної сумісності
    p.add_argument("--mode",   default=None, choices=["SHORT", "FULL"],
                   help=argparse.SUPPRESS)
    args = p.parse_args()

    # --mode перекриває --type якщо передано явно
    object_name = args.object
    mode        = args.mode or "SHORT"

    # ── Перевірка наявності папки об'єкта ──
    obj_dir = OBJECTS_DIR / object_name
    if not obj_dir.exists():
        sys.exit(f"❌ Папка об'єкта не знайдена: {obj_dir}")

    print(f"\n  Об'єкт: {object_name}  |  Режим: {mode}")

    # ── Завантаження даних ──
    subject = load_subject_yaml(object_name)
    if not subject:
        print(f"  ⚠ YAML-картка об'єкта не знайдена. "
              f"Перевірте wiki/objects/ у папці об'єкта.")
        subject = {"name": object_name}

    print(f"  Площа: {subject.get('Area', '—')} м²  |  "
          f"Клас: {subject.get('Building_Class', '—')}")

    media = find_media(object_name)
    print(f"  Фото: {len(media['photos'])}  |  Плани: {len(media['floorplans'])}")

    if media["photos"]:
        print("  Фото:", [p.name for p in media["photos"]])
    if not media["photos"]:
        print("  ⚠ Фото не знайдено в папці об'єкта.")
    if not media["floorplans"]:
        print("  ⚠ Плани поверхів не знайдено. "
              "Слайд буде створено з плейсхолдером.")

    # ── Метрики (для FULL) ──
    metrics = None
    if mode == "FULL":
        metrics = load_metrics(object_name)
        if metrics:
            b1 = metrics.get("b1", {})
            b2 = metrics.get("b2", {})
            print(f"  Метрики: rent=${b1.get('market_rent','—')}/м², "
                  f"value=${b2.get('value','—'):,}" if b2.get("value") else
                  f"  Метрики: {metrics.get('date','—')}")
        else:
            print("  ⚠ Метрики не знайдено. У FULL-режимі фінансові слайди "
                  "матимуть плейсхолдери.\n"
                  "  → Запустіть: uv run agents/valuation_report.py "
                  f"--object-dir {object_name} --excel-only --no-audit "
                  "--no-interactive --cap-rate 0.11 --discount 0.08")

    # ── Папка презентації ──
    pres_dir = obj_dir / "wiki" / "presentation"
    pres_dir.mkdir(parents=True, exist_ok=True)

    # ── Генерація Marp ──
    frontmatter = _frontmatter(object_name, mode)

    if mode == "SHORT":
        body = build_short(subject, media, object_name, pres_dir)
    else:
        body = build_full(subject, media, metrics, object_name, pres_dir)

    content = frontmatter + "\n\n---\n\n" + body + "\n"

    # ── Збереження ──
    file_name = f"{_safe_name(object_name)}_{mode}_Presentation.md"
    out_path  = pres_dir / file_name
    out_path.write_text(content, encoding="utf-8")

    print(f"\n  ✅ Презентацію збережено: {out_path}")
    print(f"     Слайдів: {content.count('---') - 1}")
    print(f"\n  Відкрити в VS Code з розширенням Marp for VS Code")
    print(f"  або: npx @marp-team/marp-cli \"{out_path}\" --pdf")
    print(f"\nPresentation system is ready")


if __name__ == "__main__":
    main()
