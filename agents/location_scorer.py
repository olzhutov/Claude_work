"""
Агент оценки локации для объектов коммерческой недвижимости.

Принимает адрес и тип объекта, автономно собирает данные (веб + знания Claude)
и возвращает структурированную оценку по 6 факторам с отчётом на украинском языке.
"""

import argparse
import json
import sys
import urllib.error
import urllib.request
from typing import Any, Dict, List, Optional, Tuple

import anthropic

from agents import config

# ---------------------------------------------------------------------------
# ПРОФИЛИ ВЕСОВ ПО ТИПАМ ОБЪЕКТОВ (скопировано из CLAUDE.md)
# Ключи начинающиеся с '_' — подсказки для агента, в расчёт не идут
# ---------------------------------------------------------------------------

LOCATION_PROFILES: Dict[str, Dict[str, Any]] = {

    # СКЛАД / ЛОГІСТИКА
    "склад": {
        "транспорт":      0.30,
        "інфраструктура": 0.20,
        "ринок_праці":    0.15,
        "конкуренти":     0.10,
        "потенціал":      0.10,
        "безпека":        0.15,
        "_транспорт_hints": ["траса_М_або_Н", "відстань_до_КПП", "ж/д_гілка",
                             "під'їзд_фур_22т", "розв'язка_поряд"],
    },

    # ОФІС / БЦ
    "офіс": {
        "транспорт":      0.25,
        "ринок_праці":    0.25,
        "інфраструктура": 0.10,
        "конкуренти":     0.15,
        "потенціал":      0.10,
        "безпека":        0.15,
        "_транспорт_hints": ["метро_пішки", "маршрутки", "парковка",
                             "центр_чи_периферія", "престижність_адреси"],
    },

    # РИТЕЙЛ / ТОРГІВЛЯ
    "ритейл": {
        "трафік":         0.30,
        "демографія":     0.20,
        "транспорт":      0.15,
        "конкуренти":     0.15,
        "потенціал":      0.05,
        "безпека":        0.15,
        "_трафік_hints":  ["пішоходи_на_добу", "якірні_орендарі_поряд",
                           "видимість_з_дороги", "ЖК_поряд", "ринок_поряд"],
    },

    # ВИРОБНИЦТВО / ЗАВОД
    "виробництво": {
        "інфраструктура": 0.30,
        "транспорт":      0.25,
        "ринок_праці":    0.20,
        "конкуренти":     0.05,
        "потенціал":      0.05,
        "безпека":        0.15,
        "_інфра_hints":   ["електро_кВт_доступно", "газ_тиск", "вода_м3_добу",
                           "ж/д_гілка", "промзона_чи_ні"],
    },

    # ГОТЕЛЬНО-РЕСТОРАННИЙ КОМПЛЕКС
    "готель": {
        "трафік":         0.25,
        "транспорт":      0.20,
        "демографія":     0.20,
        "конкуренти":     0.15,
        "потенціал":      0.05,
        "безпека":        0.15,
        "_трафік_hints":  ["центр_міста", "вокзал_хв", "туристичні_атракції",
                           "ділові_центри_поряд"],
    },

    # БАЗОВИЙ ПРОФІЛЬ (якщо тип невідомий)
    "базовий": {
        "транспорт":      0.20,
        "демографія":     0.20,
        "інфраструктура": 0.15,
        "конкуренти":     0.15,
        "потенціал":      0.15,
        "безпека":        0.15,
    },
}

# ---------------------------------------------------------------------------
# НОРМАЛИЗАЦИЯ ТИПОВ ОБЪЕКТОВ
# ---------------------------------------------------------------------------

TYPE_ALIASES: Dict[str, str] = {
    # Английские варианты
    "warehouse": "склад",
    "logistics": "склад",
    "storage": "склад",
    "retail": "ритейл",
    "shop": "ритейл",
    "shopping": "ритейл",
    "office": "офіс",
    "business_center": "офіс",
    "production": "виробництво",
    "manufacturing": "виробництво",
    "factory": "виробництво",
    "industrial": "виробництво",
    "hotel": "готель",
    "restaurant": "готель",
    "hospitality": "готель",
    # Украинские варианты (pass-through)
    "склад": "склад",
    "офіс": "офіс",
    "ритейл": "ритейл",
    "виробництво": "виробництво",
    "готель": "готель",
    "базовий": "базовий",
    # Русские варианты
    "склад_ru": "склад",
    "офис": "офіс",
    "ритейл_ru": "ритейл",
    "производство": "виробництво",
    "гостиница": "готель",
}

# ---------------------------------------------------------------------------
# СТАТУСЫ ПО ПОРОГОВЫМ ЗНАЧЕНИЯМ
# ---------------------------------------------------------------------------

# (минимальный балл, статус, emoji)
STATUS_MAP: List[Tuple[float, str, str]] = [
    (8.0, "відмінна", "✅"),
    (6.0, "приваблива", "🟡"),
    (4.0, "середня", "🟠"),
    (0.0, "слабка", "🔴"),
]

# Все типы для сравнительного анализа (кроме базового)
ALL_TYPES = ["склад", "офіс", "ритейл", "виробництво", "готель"]


# ---------------------------------------------------------------------------
# ТИПИЗИРОВАННЫЙ РЕЗУЛЬТАТ
# ---------------------------------------------------------------------------

class LocationScore(dict):
    """
    Результат оценки локации.

    Поля:
        address (str): адрес объекта
        property_type (str): нормализованный тип ("склад", "офіс", ...)
        total_score (float): итоговый взвешенный балл 0-10
        factor_scores (Dict[str, float]): балл по каждому фактору
        factor_explanations (Dict[str, str]): пояснение по каждому фактору
        factor_sources (Dict[str, str]): источник данных по каждому фактору
        status (str): "відмінна" / "приваблива" / "середня" / "слабка"
        recommendation (str): краткая рекомендация
        strengths (List[str]): сильные стороны
        risks (List[str]): риски
        sources (List[str]): использованные источники
        report (str): полный форматированный отчёт на украинском
    """
    pass


# ---------------------------------------------------------------------------
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ---------------------------------------------------------------------------

def _fetch_url(url: str, timeout: int = 10) -> Optional[str]:
    """
    Получает текст страницы по URL. При ошибке возвращает None.
    Обрезает до 3000 символов чтобы не перегружать промпт.
    """
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; CRE-Analyzer/1.0)"}
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
            # Декодируем, игнорируя ошибки
            text = raw.decode("utf-8", errors="ignore")
            # Убираем HTML-теги грубым образом
            import re
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text).strip()
            return text[:3000]
    except Exception:
        return None


def _extract_city(address: str) -> str:
    """
    Извлекает название города из адреса.
    Ищет последнее слово перед запятой или конец строки.
    """
    # Пробуем распознать по типичным паттернам украинских адресов:
    # "вул. Xxx, N, Місто" или "Місто, вул. Xxx, N"
    parts = [p.strip() for p in address.split(",")]
    # Последняя часть чаще всего — город
    for part in reversed(parts):
        # Убираем распространённые слова
        cleaned = part.replace("м.", "").replace("м ", "").strip()
        if cleaned and len(cleaned) > 2:
            return cleaned
    return address


def _gather_web_data(address: str) -> str:
    """
    Собирает данные о локации из открытых веб-источников.
    Игнорирует ошибки — если сайт недоступен, продолжаем работу.
    Возвращает строку со всеми полученными данными (или пустую строку).
    """
    city = _extract_city(address)
    snippets: List[str] = []

    # Wikipedia: информация о городе (население, описание)
    city_slug = city.replace(" ", "_")
    wiki_url = f"https://en.wikipedia.org/wiki/{city_slug}"
    wiki_data = _fetch_url(wiki_url)
    if wiki_data:
        snippets.append(f"[Wikipedia про {city}]:\n{wiki_data[:2000]}")

    # Украиноязычная Wikipedia
    wiki_ua_url = f"https://uk.wikipedia.org/wiki/{city_slug}"
    wiki_ua_data = _fetch_url(wiki_ua_url)
    if wiki_ua_data:
        snippets.append(f"[Вікіпедія про {city}]:\n{wiki_ua_data[:2000]}")

    return "\n\n".join(snippets) if snippets else ""


def _format_profile_for_prompt(profile: Dict[str, Any]) -> Tuple[str, str]:
    """
    Форматирует профиль весов для промпта.
    Возвращает (weights_str, hints_str).
    """
    weights_lines: List[str] = []
    hints_lines: List[str] = []

    for key, value in profile.items():
        if key.startswith("_"):
            # Это подсказка
            factor_name = key[1:].replace("_hints", "")
            hints_lines.append(f"  {factor_name}: {', '.join(value)}")
        else:
            weights_lines.append(f"  {key}: {value:.0%}")

    return "\n".join(weights_lines), "\n".join(hints_lines) if hints_lines else "  (загальні)"


def _ask_claude(
    address: str,
    property_type: str,
    profile: Dict[str, Any],
    web_data: str,
) -> Dict[str, Any]:
    """
    Отправляет запрос в Claude для оценки локации.
    Возвращает словарь с factor_scores, explanations и т.д.
    """
    client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)

    weights_str, hints_str = _format_profile_for_prompt(profile)

    # Собираем список факторов (без ключей начинающихся с _)
    factors = [k for k in profile.keys() if not k.startswith("_")]
    factors_json = "{" + ", ".join(f'"{f}": 0.0' for f in factors) + "}"

    web_section = ""
    if web_data:
        web_section = f"\n[ДАНІ З ВЕБ-ПОШУКУ]:\n{web_data}\n"

    prompt = f"""Ти — експерт з аналізу локацій комерційної нерухомості в Україні.

Оціни локацію для об'єкта нерухомості.

АДРЕСА: {address}
ТИП ОБ'ЄКТА: {property_type}

ПРОФІЛЬ ОЦІНКИ (фактор: вага в загальній оцінці):
{weights_str}

ОРІЄНТИРИ ДЛЯ ОЦІНКИ:
{hints_str}
{web_section}
ЗАВДАННЯ:
Оціни кожен фактор від 0 до 10, використовуючи:
1. Свої знання про Україну, ринок нерухомості, транспортну мережу та воєнний стан
2. Дані з веб-пошуку (якщо є вище)
3. Логічні припущення на основі типу об'єкта та адреси

ВАЖЛИВО для фактору "безпека":
- Враховуй відстань від лінії фронту станом на 2025-2026 рр.
- Київ, Львів, Ужгород, Дніпро, Харків, Запоріжжя — різні рівні безпеки
- Східні міста (Харків, Запоріжжя) мають підвищений ризик через близькість фронту

Поверни ТІЛЬКИ валідний JSON (без жодних пояснень до або після):
{{
  "factor_scores": {factors_json},
  "factor_explanations": {{{", ".join(f'"{f}": "1 речення пояснення"' for f in factors)}}},
  "factor_sources": {{{", ".join(f'"{f}": "джерело даних"' for f in factors)}}},
  "strengths": ["сильна сторона 1", "сильна сторона 2", "сильна сторона 3"],
  "risks": ["⚠️ ризик 1", "⚠️ ризик 2"],
  "recommendation": "2-3 речення загального висновку та рекомендації"
}}"""

    message = client.messages.create(
        model=config.CLAUDE_MODEL,
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    )

    # Извлекаем текст ответа
    response_text = message.content[0].text.strip()

    # Убираем markdown-обёртки если есть (как в extractor.py)
    if response_text.startswith("```"):
        lines = response_text.split("\n")
        # Убираем первую и последнюю строки (```json ... ```)
        response_text = "\n".join(lines[1:-1])

    return json.loads(response_text)


def _calculate_total(
    factor_scores: Dict[str, float],
    profile: Dict[str, Any],
) -> float:
    """
    Вычисляет взвешенную итоговую оценку.
    Ключи начинающиеся с '_' пропускаются (подсказки, не веса).
    Если фактор не найден в оценке — используем 5.0 как нейтральный дефолт.
    """
    total = 0.0
    for key, value in profile.items():
        if key.startswith("_"):
            continue
        score = factor_scores.get(key, 5.0)
        # Ограничиваем диапазон 0-10
        score = max(0.0, min(10.0, float(score)))
        total += score * float(value)
    return round(total, 2)


def _get_status(score: float) -> Tuple[str, str]:
    """
    Определяет статус и emoji по итоговому баллу.
    Возвращает (статус, emoji).
    """
    for threshold, status, emoji in STATUS_MAP:
        if score >= threshold:
            return status, emoji
    return "слабка", "🔴"


def _build_report(
    address: str,
    property_type: str,
    total_score: float,
    factor_scores: Dict[str, float],
    factor_explanations: Dict[str, str],
    factor_sources: Dict[str, str],
    profile: Dict[str, Any],
    status: str,
    emoji: str,
    recommendation: str,
    strengths: List[str],
    risks: List[str],
    sources: List[str],
) -> str:
    """
    Формирует форматированный украиноязычный отчёт.
    """
    lines: List[str] = []
    separator = "━" * 55

    lines.append("ОЦІНКА ЛОКАЦІЇ")
    lines.append(f"Адреса: {address}")
    lines.append(f"Тип об'єкта: {property_type.upper()}")
    lines.append(separator)
    lines.append(f"Загальна оцінка:  {total_score:.1f} / 10  {emoji} {status.upper()}")
    lines.append("")

    # Факторы
    for key, weight in profile.items():
        if key.startswith("_"):
            continue
        score = factor_scores.get(key, 5.0)
        explanation = factor_explanations.get(key, "")
        source = factor_sources.get(key, "власні знання")
        weight_pct = int(round(float(weight) * 100))
        lines.append(f"{key.capitalize()} ({weight_pct}%):  {score:.1f}/10  — {explanation}")
        lines.append(f"    Джерело: {source}")

    lines.append("")
    lines.append(f"ВИСНОВОК: {recommendation}")
    lines.append("")

    if strengths:
        lines.append("СИЛЬНІ СТОРОНИ:")
        for s in strengths:
            lines.append(f"  ✓ {s}")
        lines.append("")

    if risks:
        lines.append("РИЗИКИ:")
        for r in risks:
            # Убедимся что ⚠️ на месте
            if not r.startswith("⚠️"):
                r = f"⚠️ {r}"
            lines.append(f"  {r}")
        lines.append("")

    if sources:
        lines.append(f"ДЖЕРЕЛА: {', '.join(sources)}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# ПУБЛИЧНЫЕ ФУНКЦИИ
# ---------------------------------------------------------------------------

def score_location(
    address: str,
    property_type: Optional[str] = None,
) -> LocationScore:
    """
    Главная функция оценки локации.

    Args:
        address: Адрес объекта (например: "вул. Таращанська, 191-А, Біла Церква")
        property_type: Тип объекта. Если None — использует базовый профиль.
                      Поддерживаемые значения: "склад", "офіс", "ритейл",
                      "виробництво", "готель" (и английские аналоги).

    Returns:
        LocationScore с полной оценкой и форматированным отчётом.
    """
    # Нормализуем тип
    if property_type is None:
        normalized_type = "базовий"
    else:
        normalized_type = TYPE_ALIASES.get(property_type.lower().strip(), "базовий")

    profile = LOCATION_PROFILES[normalized_type]

    # Этап 1: Собираем веб-данные
    print(f"  Збираємо дані для: {address} [{normalized_type}]...", file=sys.stderr)
    web_data = _gather_web_data(address)

    # Этап 2: Анализ Claude
    print("  Аналізуємо локацію...", file=sys.stderr)
    claude_result = _ask_claude(address, normalized_type, profile, web_data)

    # Извлекаем результаты
    factor_scores: Dict[str, float] = claude_result.get("factor_scores", {})
    factor_explanations: Dict[str, str] = claude_result.get("factor_explanations", {})
    factor_sources: Dict[str, str] = claude_result.get("factor_sources", {})
    strengths: List[str] = claude_result.get("strengths", [])
    risks: List[str] = claude_result.get("risks", [])
    recommendation: str = claude_result.get("recommendation", "")

    # Рассчитываем итог
    total_score = _calculate_total(factor_scores, profile)
    status, emoji = _get_status(total_score)

    # Формируем список источников
    sources: List[str] = ["Claude AI (власні знання про Україну)"]
    if web_data:
        sources.append("Wikipedia")

    # Строим отчёт
    report = _build_report(
        address=address,
        property_type=normalized_type,
        total_score=total_score,
        factor_scores=factor_scores,
        factor_explanations=factor_explanations,
        factor_sources=factor_sources,
        profile=profile,
        status=status,
        emoji=emoji,
        recommendation=recommendation,
        strengths=strengths,
        risks=risks,
        sources=sources,
    )

    result = LocationScore(
        address=address,
        property_type=normalized_type,
        total_score=total_score,
        factor_scores=factor_scores,
        factor_explanations=factor_explanations,
        factor_sources=factor_sources,
        status=status,
        recommendation=recommendation,
        strengths=strengths,
        risks=risks,
        sources=sources,
        report=report,
    )

    return result


def compare_all_types(address: str) -> Dict[str, LocationScore]:
    """
    Оценивает один адрес для всех 5 типов объектов.
    Возвращает словарь {тип: LocationScore}.
    """
    results: Dict[str, LocationScore] = {}
    for prop_type in ALL_TYPES:
        print(f"\n[{prop_type.upper()}]", file=sys.stderr)
        results[prop_type] = score_location(address, prop_type)
    return results


def format_comparison_table(scores: Dict[str, LocationScore]) -> str:
    """
    Форматирует сравнительную таблицу для всех типов объектов.
    """
    lines: List[str] = []
    separator = "━" * 55

    # Получаем адрес из первого результата
    address = next(iter(scores.values()))["address"] if scores else ""

    lines.append("ПОРІВНЯЛЬНА ОЦІНКА ЛОКАЦІЇ")
    lines.append(f"Адреса: {address}")
    lines.append(separator)
    lines.append("Та сама адреса, різні призначення:")
    lines.append("")

    # Сортируем по убыванию балла
    sorted_scores = sorted(
        scores.items(),
        key=lambda x: x[1]["total_score"],
        reverse=True
    )

    for prop_type, score in sorted_scores:
        total = score["total_score"]
        status = score["status"]
        _, emoji = _get_status(total)
        type_label = prop_type.capitalize()
        lines.append(f"  {type_label:<16} {total:.1f} / 10  {emoji} {status}")

    lines.append("")

    # Рекомендованные типы (балл >= 6)
    good_types = [
        f"{t.capitalize()} ({s['total_score']:.1f})"
        for t, s in sorted_scores
        if s["total_score"] >= 6.0
    ]
    bad_types = [
        f"{t.capitalize()} ({s['total_score']:.1f})"
        for t, s in sorted_scores
        if s["total_score"] < 4.0
    ]

    if good_types:
        lines.append(f"→ Оптимально для: {', '.join(good_types)}")
    if bad_types:
        lines.append(f"→ Не рекомендується: {', '.join(bad_types)}")

    return "\n".join(lines)


def print_report(score: LocationScore) -> None:
    """Выводит отчёт в stdout."""
    print(score["report"])


# ---------------------------------------------------------------------------
# CLI-ИНТЕРФЕЙС
# ---------------------------------------------------------------------------

def main() -> None:
    """Точка входа для CLI-запуска."""
    parser = argparse.ArgumentParser(
        description="Оцінка локації для об'єктів комерційної нерухомості",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Приклади:
  python3 agents/location_scorer.py "вул. Таращанська, 191-А, Біла Церква" --type ритейл
  python3 agents/location_scorer.py "вул. Таращанська, 191-А, Біла Церква" --type склад
  python3 agents/location_scorer.py "вул. Таращанська, 191-А, Біла Церква"
  python3 agents/location_scorer.py "вул. Володимирська, 8, Київ" --type офіс

Типи об'єктів: склад, офіс, ритейл, виробництво, готель
  (або: warehouse, office, retail, production, hotel)
        """,
    )
    parser.add_argument(
        "address",
        help="Адреса об'єкта (наприклад: 'вул. Таращанська, 191-А, Біла Церква')"
    )
    parser.add_argument(
        "--type", "-t",
        dest="property_type",
        default=None,
        help="Тип об'єкта. Якщо не вказано — показати порівняння для всіх типів."
    )

    args = parser.parse_args()

    if args.property_type is None:
        # Режим сравнения всех типов
        print(f"\nОцінюємо локацію для всіх типів об'єктів...", file=sys.stderr)
        print(f"Адреса: {args.address}\n", file=sys.stderr)
        scores = compare_all_types(args.address)
        print("\n" + format_comparison_table(scores))
        # Дополнительно выводим полный отчёт для лучшего типа
        best_type = max(scores.items(), key=lambda x: x[1]["total_score"])
        print(f"\n{'━' * 55}")
        print(f"ДЕТАЛЬНИЙ ЗВІТ ДЛЯ НАЙКРАЩОГО ТИПУ ({best_type[0].upper()}):")
        print("━" * 55)
        print_report(best_type[1])
    else:
        # Оценка конкретного типа
        score = score_location(args.address, args.property_type)
        print()
        print_report(score)


if __name__ == "__main__":
    main()
