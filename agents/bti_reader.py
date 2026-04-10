"""
Агент читання документів БТІ.
Знаходить і аналізує технічні паспорти, Справи та Витяги ДРРП.
Пріоритет джерел: експлікація > Витяг ДРРП > план/креслення.

Публічний API:
    read_bti_areas(folder_path) -> BtiAreaData
"""

import base64
import json
import re
import sys
from pathlib import Path
from typing import Optional

from anthropic import Anthropic

from config import ANTHROPIC_API_KEY, CLAUDE_MODEL
from schemas import BtiAreaData, BtiRoom, BtiSprava


# Ключові слова для розпізнавання БТІ-документів
_BTI_KEYWORDS = [
    "бті", "бти", "технічний паспорт", "технический паспорт",
    "справа №", "справа_", "спр.", "техпаспорт",
    "інвентаризац", "инвентаризац", "план поверху", "поверховий план",
    "план поверха", "floor plan",
]

# Ключові слова для Витягу ДРРП
_DRRP_KEYWORDS = [
    "дррп", "держреєстр", "витяг", "реєстр речових",
    "реестр прав", "дрп", "nais",
]


def read_bti_areas(folder_path: str) -> BtiAreaData:
    """
    Читає всі БТІ-документи в папці об'єкта та повертає верифіковані площі.

    Args:
        folder_path: Шлях до папки об'єкта (містить raw/ директорію)

    Returns:
        BtiAreaData: Верифіковані дані про площі або порожній результат
                     якщо документів БТІ не знайдено
    """
    folder = Path(folder_path)
    raw_dir = folder / "raw"

    if not raw_dir.exists():
        return _empty_result("raw/ директорія відсутня")

    # Знаходимо всі файли в raw/documents і raw/plans
    all_files = []
    for subdir in ["documents", "plans"]:
        subdir_path = raw_dir / subdir
        if subdir_path.exists():
            all_files.extend(subdir_path.iterdir())

    if not all_files:
        return _empty_result("файли в raw/ відсутні")

    # Розбиваємо на БТІ-документи, ДРРП та решту
    bti_files, drrp_files = _classify_files(all_files)

    if not bti_files and not drrp_files:
        return _empty_result("документів БТІ або ДРРП не знайдено")

    client = Anthropic()
    source_files = []

    # Читаємо БТІ-документи → збираємо Справи
    spravy: list = []
    all_discrepancies: list = []
    bti_date: Optional[str] = None
    overall_confidence = "low"

    for file_path in bti_files:
        source_files.append(file_path.name)
        result = _extract_from_bti_document(client, file_path)
        if result:
            if result.get("spravy"):
                spravy.extend(result["spravy"])
            if result.get("discrepancies"):
                all_discrepancies.extend(result["discrepancies"])
            if result.get("bti_date") and not bti_date:
                bti_date = result["bti_date"]
            if result.get("confidence") == "high":
                overall_confidence = "high"
            elif result.get("confidence") == "medium" and overall_confidence == "low":
                overall_confidence = "medium"

    # Читаємо ДРРП
    drrp_area: Optional[float] = None
    for file_path in drrp_files:
        source_files.append(file_path.name)
        drrp_result = _extract_drrp_area(client, file_path)
        if drrp_result is not None:
            drrp_area = drrp_result
            break  # беремо перший знайдений ДРРП

    # Якщо БТІ-документів немає але є ДРРП
    if not spravy and drrp_area:
        return BtiAreaData(
            total_area=drrp_area,
            usable_area=drrp_area,
            auxiliary_area=0.0,
            spravy=[],
            drrp_area=drrp_area,
            drrp_match=True,
            discrepancies=[],
            source_files=source_files,
            bti_date=None,
            confidence="medium",
            notes="Дані тільки з ДРРП, документів БТІ не знайдено",
        )

    # Рахуємо підсумки по всіх Справах
    total_area = sum(s.get("total_area", 0) for s in spravy)
    usable_area = sum(s.get("usable_area", 0) for s in spravy)
    auxiliary_area = sum(s.get("auxiliary_area", 0) for s in spravy)

    # Порівнюємо з ДРРП
    drrp_match: Optional[bool] = None
    if drrp_area is not None and total_area > 0:
        drrp_match = abs(total_area - drrp_area) <= 0.5

    # Формуємо примітки
    notes_parts = []
    if spravy:
        notes_parts.append(f"Знайдено Справ: {len(spravy)}")
    if drrp_area:
        diff = total_area - drrp_area
        sign = "+" if diff >= 0 else ""
        notes_parts.append(f"ДРРП: {drrp_area} м² ({sign}{diff:.1f} м²)")
    if all_discrepancies:
        notes_parts.append(f"Розбіжності план≠експл: {len(all_discrepancies)} шт.")

    # Зберігаємо результат в extracted/bti_areas.json
    result = BtiAreaData(
        total_area=round(total_area, 1),
        usable_area=round(usable_area, 1),
        auxiliary_area=round(auxiliary_area, 1),
        spravy=spravy,
        drrp_area=drrp_area,
        drrp_match=drrp_match,
        discrepancies=all_discrepancies,
        source_files=source_files,
        bti_date=bti_date,
        confidence=overall_confidence,
        notes=" | ".join(notes_parts) if notes_parts else "",
    )

    _save_result(folder, result)
    return result


def _classify_files(files: list) -> tuple:
    """
    Розбиває список файлів на БТІ-документи та ДРРП.
    Повертає (bti_files, drrp_files).
    """
    supported = {".pdf", ".jpg", ".jpeg", ".png", ".webp", ".txt", ".docx"}
    bti_files = []
    drrp_files = []

    for f in files:
        if f.suffix.lower() not in supported:
            continue

        name_lower = f.name.lower()

        # Перевіряємо по імені файлу
        is_drrp = any(kw in name_lower for kw in _DRRP_KEYWORDS)
        is_bti = any(kw in name_lower for kw in _BTI_KEYWORDS)

        if is_drrp:
            drrp_files.append(f)
        elif is_bti:
            bti_files.append(f)
        else:
            # Невідомий файл — додаємо до БТІ для перевірки
            # (агент сам визначить чи є там БТІ-дані)
            bti_files.append(f)

    return bti_files, drrp_files


def _extract_from_bti_document(client: Anthropic, file_path: Path) -> Optional[dict]:
    """
    Надсилає один документ у Claude API з промптом для читання БТІ.
    Повертає dict зі Справами або None якщо документ не є БТІ.
    """
    try:
        content = _read_file_for_api(file_path)
    except Exception as e:
        print(f"  ⚠ Не вдалося прочитати {file_path.name}: {e}", file=sys.stderr)
        return None

    system_prompt = """Ти — фахівець з читання технічних паспортів БТІ (Бюро технічної інвентаризації).

ПРАВИЛО ПРІОРИТЕТУ ДАНИХ:
1. Таблиця-експлікація (розшифровка) — НАЙВИЩИЙ пріоритет. Якщо є — беремо звідти.
2. Числові значення на кресленні/плані — НИЖЧИЙ пріоритет.
3. Якщо значення на плані ≠ значення в експлікації → берємо з ЕКСПЛІКАЦІЇ і відзначаємо розбіжність.

Твоя задача:
- Знайти всі таблиці-експлікації (розшифровки площ) в документі
- Витягнути для кожної Справи: список приміщень з площами та типами
- Визначити загальну, корисну (основну) та допоміжну площу
- Знайти дату технічного паспорту якщо є
- Якщо документ НЕ є БТІ — повернути {"is_bti": false}

ТИПИ ПЛОЩ:
- "основна" — житлові кімнати, офісні приміщення, торгові зали, склади
- "допоміжна" — коридори, санвузли, сходи, тамбури, підсобні

ПОВЕРТАЙ ТІЛЬКИ JSON БЕЗ ПОЯСНЕНЬ."""

    user_message = f"""Прочитай документ БТІ та витягни дані про площі.

Документ: {file_path.name}

{content}

Поверни JSON:
{{
  "is_bti": true,
  "bti_date": "рік або null",
  "confidence": "high/medium/low",
  "spravy": [
    {{
      "name": "Справа №1 — МЗК",
      "total_area": 58.1,
      "usable_area": 45.2,
      "auxiliary_area": 12.9,
      "rooms": [
        {{"id": "1", "name": "кімната", "area": 25.3, "area_type": "основна", "source": "Справа №1 — МЗК"}},
        {{"id": "2", "name": "коридор", "area": 8.1, "area_type": "допоміжна", "source": "Справа №1 — МЗК"}}
      ]
    }}
  ],
  "discrepancies": [
    {{"room": "№5", "plan_value": 18.0, "expl_value": 18.3, "source": "Справа №2"}}
  ]
}}"""

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=4000,
        system=system_prompt,
        messages=[{"role": "user", "content": _build_message_content(file_path, user_message)}],
    )

    response_text = response.content[0].text.strip()
    response_text = _strip_markdown(response_text)

    try:
        data = json.loads(response_text)
    except json.JSONDecodeError:
        return None

    if not data.get("is_bti", True):
        return None

    return data


def _extract_drrp_area(client: Anthropic, file_path: Path) -> Optional[float]:
    """
    Читає Витяг ДРРП і повертає площу об'єкта.
    Повертає None якщо не вдалося визначити.
    """
    try:
        content = _read_file_for_api(file_path)
    except Exception:
        return None

    system_prompt = """Ти читаєш Витяг з Державного реєстру речових прав (ДРРП).
Твоя задача: знайти загальну площу об'єкта нерухомості.
Повертай ТІЛЬКИ JSON без пояснень."""

    user_message = f"""Знайди площу об'єкта у Витязі ДРРП.

{content}

Поверни JSON:
{{"area": 681.9, "unit": "м²"}}

Якщо площу не знайдено: {{"area": null}}"""

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=200,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )

    response_text = _strip_markdown(response.content[0].text.strip())

    try:
        data = json.loads(response_text)
        return float(data["area"]) if data.get("area") else None
    except (json.JSONDecodeError, TypeError, ValueError):
        return None


def _read_file_for_api(file_path: Path) -> str:
    """
    Читає файл і повертає вміст у форматі для Claude API.
    Текстові файли → рядок, зображення → base64 опис.
    """
    suffix = file_path.suffix.lower()

    if suffix in {".txt", ".md"}:
        return file_path.read_text(encoding="utf-8")

    elif suffix == ".pdf":
        try:
            import PyPDF2
            pages = []
            with open(file_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    pages.append(page.extract_text())
            return "\n".join(pages)
        except ImportError:
            raise ImportError("Встановіть PyPDF2: pip3 install PyPDF2")

    elif suffix in {".jpg", ".jpeg", ".png", ".gif", ".webp"}:
        with open(file_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")

    else:
        raise ValueError(f"Непідтримуваний формат: {suffix}")


def _build_message_content(file_path: Path, text_message: str):
    """
    Будує content для Claude API.
    Для зображень використовує vision (base64), для тексту — текстовий блок.
    """
    suffix = file_path.suffix.lower()

    if suffix in {".jpg", ".jpeg", ".png", ".webp"}:
        # Читаємо зображення як base64 для vision API
        with open(file_path, "rb") as f:
            image_data = base64.b64encode(f.read()).decode("utf-8")

        media_map = {
            ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".png": "image/png", ".webp": "image/webp",
        }
        media_type = media_map.get(suffix, "image/jpeg")

        return [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": media_type,
                    "data": image_data,
                },
            },
            {"type": "text", "text": text_message},
        ]
    else:
        return text_message


def _strip_markdown(text: str) -> str:
    """Видаляє ```json ... ``` обгортку якщо є."""
    if text.startswith("```"):
        text = re.sub(r"^```[a-z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
    return text.strip()


def _empty_result(reason: str) -> BtiAreaData:
    """Повертає порожній BtiAreaData з причиною."""
    return BtiAreaData(
        total_area=0.0,
        usable_area=0.0,
        auxiliary_area=0.0,
        spravy=[],
        drrp_area=None,
        drrp_match=None,
        discrepancies=[],
        source_files=[],
        bti_date=None,
        confidence="low",
        notes=reason,
    )


def _save_result(folder: Path, result: BtiAreaData) -> None:
    """Зберігає результат в extracted/bti_areas.json."""
    extracted_dir = folder / "extracted"
    extracted_dir.mkdir(parents=True, exist_ok=True)
    output_file = extracted_dir / "bti_areas.json"
    output_file.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def format_brief_section(bti: BtiAreaData) -> str:
    """
    Форматує дані БТІ для розділу «Технічні характеристики» Інформаційної справки.
    Формат A (короткий підсумок).
    """
    if not bti["total_area"]:
        return ""

    lines = ["Площі об'єкта (за даними БТІ):"]
    lines.append(f"— Загальна площа: {bti['total_area']} м²")

    if bti["usable_area"]:
        lines.append(f"— Корисна площа: {bti['usable_area']} м²")
    if bti["auxiliary_area"]:
        lines.append(f"— Допоміжна площа: {bti['auxiliary_area']} м²")

    if bti["source_files"]:
        sources = ", ".join(bti["source_files"])
        lines.append(f"\nДжерело: Технічний паспорт БТІ ({sources})")

    if bti["bti_date"]:
        lines.append(f"Дата документів: {bti['bti_date']}")

    if bti["drrp_area"] is not None:
        if bti["drrp_match"]:
            lines.append(f"Витяг ДРРП: {bti['drrp_area']} м² — збігається з БТІ ✅")
        else:
            diff = bti["total_area"] - bti["drrp_area"]
            sign = "+" if diff >= 0 else ""
            lines.append(
                f"⚠️ Розбіжність з ДРРП: {sign}{diff:.1f} м² "
                f"(площа за ДРРП: {bti['drrp_area']} м²)"
            )

    return "\n".join(lines)


if __name__ == "__main__":
    # CLI: python3 bti_reader.py <folder_path>
    if len(sys.argv) < 2:
        print("Використання: python3 bti_reader.py <папка_об'єкта>")
        sys.exit(1)

    folder = sys.argv[1]
    print(f"Читаю БТІ-документи: {folder}")
    print()

    result = read_bti_areas(folder)

    print("=== РЕЗУЛЬТАТ ===")
    print(f"Загальна площа:   {result['total_area']} м²")
    print(f"Корисна площа:    {result['usable_area']} м²")
    print(f"Допоміжна площа:  {result['auxiliary_area']} м²")
    print(f"Кількість Справ:  {len(result['spravy'])}")
    print(f"Впевненість:      {result['confidence']}")

    if result["drrp_area"]:
        match_str = "✅ збігається" if result["drrp_match"] else "⚠️ розбіжність"
        print(f"ДРРП:             {result['drrp_area']} м² ({match_str})")

    if result["discrepancies"]:
        print(f"\nРозбіжності план≠експлікація ({len(result['discrepancies'])}):")
        for d in result["discrepancies"]:
            print(f"  {d.get('room')}: план={d.get('plan_value')}, експл.={d.get('expl_value')}")

    if result["notes"]:
        print(f"\nПримітки: {result['notes']}")

    print()
    print("=== ДЛЯ ІНФОРМАЦІЙНОЇ СПРАВКИ ===")
    print(format_brief_section(result))

    if result["total_area"] > 0:
        extracted_path = Path(folder) / "extracted" / "bti_areas.json"
        print(f"\n✓ Збережено: {extracted_path}")
