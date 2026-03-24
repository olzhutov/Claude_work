"""
Агент-генератор инвестиционного меморандума.
Создаёт текстовый или PowerPoint документ в зависимости от типа и размера.
Выходной файл: output/investment_memo.txt
"""

from typing import Optional, Dict, Any
from pathlib import Path

from schemas import PropertyData, PipelineConfig


def generate_memo(
    property_data: PropertyData,
    metrics: Optional[Dict[str, Any]],
    config: PipelineConfig,
    output_dir: str = "output",
) -> str:
    """
    Генерирует инвестиционный меморандум.

    Args:
        property_data: Извлечённые данные об объекте
        metrics: Финансовые метрики (или None для перспективных объектов)
        config: Конфигурация pipeline
        output_dir: Директория для выходных файлов

    Returns:
        Путь к созданному файлу
    """

    content = _build_memo_content(property_data, metrics, config)

    output_path = Path(output_dir) / "investment_memo.txt"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")

    return str(output_path)


def _build_memo_content(
    property_data: PropertyData,
    metrics: Optional[Dict[str, Any]],
    config: PipelineConfig,
) -> str:
    """Собирает содержимое меморандума."""

    lines = []
    currency = config["currency"]
    currency_symbol = "$" if currency == "USD" else "₴"

    # Обложка
    lines.append("=" * 70)
    lines.append(f"ІНВЕСТИЦІЙНИЙ МЕМОРАНДУМ")
    lines.append("=" * 70)
    lines.append("")
    lines.append(f"Об'єкт: {property_data['property_name']}")
    lines.append(f"Тип: {property_data['property_type'].upper()}")
    lines.append(f"Місце розташування: {property_data['city']}")
    if property_data.get("address"):
        lines.append(f"Адреса: {property_data['address']}")
    lines.append("")
    lines.append("=" * 70)
    lines.append("")

    # Резюме
    lines.append("1. РЕЗЮМЕ")
    lines.append("-" * 70)
    lines.append("")
    lines.append(property_data.get("description", "Опис об'єкта недоступний."))
    lines.append("")

    # Характеристики
    lines.append("2. ХАРАКТЕРИСТИКИ ОБ'ЄКТУ")
    lines.append("-" * 70)
    lines.append("")
    lines.append(f"Валовая площа (GBA): {property_data['gba']:,.0f} м²")

    if property_data.get("gla"):
        lines.append(f"Арендуема площа (GLA): {property_data['gla']:,.0f} м²")

    if property_data.get("year_built"):
        lines.append(f"Рік побудови: {property_data['year_built']}")

    if property_data.get("condition"):
        lines.append(f"Стан: {property_data['condition']}")

    if property_data.get("land_area"):
        lines.append(f"Площа земельної ділянки: {property_data['land_area']} га")

    if property_data.get("infrastructure"):
        lines.append(f"Локація: {property_data['infrastructure']}")

    lines.append("")

    # Финансовые показатели (для доходных объектов)
    if metrics:
        lines.append("3. ФІНАНСОВІ ПОКАЗАТЕЛИ")
        lines.append("-" * 70)
        lines.append("")
        lines.append(f"Вартість об'єкту: {metrics['value']:,.0f} {currency_symbol}")
        lines.append("")

        if metrics.get("noi"):
            lines.append(f"Чистий операційний дохід (NOI): {metrics['noi']:,.0f} {currency_symbol}/рік")

        if metrics.get("egi"):
            lines.append(f"Ефективний валовий дохід (EGI): {metrics['egi']:,.0f} {currency_symbol}/рік")

        if metrics.get("cap_rate"):
            lines.append(f"Cap Rate: {metrics['cap_rate']:.2f}%")

        if metrics.get("payback_years"):
            lines.append(f"Термін окупаємості: {metrics['payback_years']:.1f} років")

        if metrics.get("price_per_unit"):
            unit_name = metrics.get("unit_name", "м²")
            lines.append(f"Вартість за одиницю: {metrics['price_per_unit']:,.0f} {currency_symbol}/{unit_name}")

        lines.append("")

        # Сценарии (если есть)
        if metrics.get("cap_rate_scenarios"):
            lines.append("4. СЦЕНАРІЇ ОЦІНКИ")
            lines.append("-" * 70)
            lines.append("")
            lines.append("Cap Rate (%)    | Вартість об'єкту ({}) |".format(currency_symbol))
            lines.append("-" * 40)

            for scenario in metrics["cap_rate_scenarios"]:
                # Сценарии могут быть dict или список
                if isinstance(scenario, dict):
                    cap_rate = scenario.get("cap_rate_pct")
                    value = scenario.get("property_value")
                elif isinstance(scenario, (list, tuple)) and len(scenario) >= 2:
                    cap_rate = scenario[0]
                    value = scenario[1]
                else:
                    continue

                if cap_rate and value:
                    lines.append(f"{cap_rate:>13.1f}  | {value:>20,.0f} |")

            lines.append("")

    else:
        # Для перспективных объектов
        lines.append("3. ОПИСАНИЕ ОБЪЕКТА")
        lines.append("-" * 70)
        lines.append("")
        lines.append(property_data.get("description", "Опис недоступний"))
        lines.append("")

        if property_data.get("infrastructure"):
            lines.append("4. ЛОКАЦІЯ І ІНФРАСТРУКТУРА")
            lines.append("-" * 70)
            lines.append("")
            lines.append(property_data["infrastructure"])
            lines.append("")

    # Закрытие
    lines.append("=" * 70)
    lines.append("Дата повідомлення: 2026")
    lines.append("=" * 70)

    return "\n".join(lines)


if __name__ == "__main__":
    from schemas import PropertyData

    test_data = PropertyData(
        property_name="Склад Тест",
        property_type="склад",
        address="вул. Тестова, 1",
        city="Київ",
        description="Сучасний складський комплекс вищого класу з інфраструктурою.",
        gba=5000,
        gla=4500,
        value=1200000,
        value_currency="USD",
        rent_rate=6,
        rent_rate_currency="USD",
        vacancy_rate=0.05,
        opex=50000,
        opex_currency="USD",
        rent_growth_rate=0.03,
        exit_cap_rate=0.09,
        hold_period=10,
        year_built=2015,
        condition="задовільний",
        land_area=None,
        infrastructure="біля МКАД, логістичний центр",
        source_file="test.txt",
        extraction_confidence="high",
        extraction_notes="",
    )

    test_metrics = {
        "value": 1200000,
        "noi": 86400,
        "egi": 180000,
        "cap_rate": 7.2,
        "payback_years": 13.9,
        "price_per_unit": 240,
        "unit_name": "м²",
        "cap_rate_scenarios": [
            {"cap_rate_pct": 6.5, "property_value": 1329231},
            {"cap_rate_pct": 7.0, "property_value": 1234286},
            {"cap_rate_pct": 8.0, "property_value": 1080000},
        ],
    }

    config = PipelineConfig(
        object_category="income",
        doc_type="memo",
        pres_format="pptx",
        exchange_rate=41.5,
        currency="USD",
    )

    output_file = generate_memo(test_data, test_metrics, config, output_dir="/tmp")
    print(f"Створен файл: {output_file}")
    print("\nВміст:")
    print(Path(output_file).read_text(encoding="utf-8"))
