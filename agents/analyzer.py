"""
Агент-анализатор: интегрирует PropertyData с cre_analyzer.calculate_metrics().
Выполняет финансовый анализ объекта и возвращает метрики.
"""

import sys
from pathlib import Path
from typing import Optional, Dict, Any

from schemas import PropertyData, PipelineConfig

# Добавляем родительскую директорию в path для импорта cre_analyzer
sys.path.insert(0, str(Path(__file__).parent.parent))
from cre_analyzer import calculate_metrics


def analyze_property(
    property_data: PropertyData, config: PipelineConfig
) -> Optional[Dict[str, Any]]:
    """
    Анализирует объект и возвращает финансовые метрики.
    Пропускает анализ для перспективных объектов (category="prospect").

    Args:
        property_data: Извлечённые данные об объекте
        config: Конфигурация pipeline

    Returns:
        Dict с метриками или None для перспективных объектов

    Raises:
        ValueError: Если недостаточно данных для анализа доходного объекта
    """

    # Для перспективных объектов анализ не требуется
    if config["object_category"] == "prospect":
        return None

    # Проверяем что есть необходимые данные для финансового анализа
    if property_data.get("rent_rate") is None:
        raise ValueError(
            "Для доходного объекта необходима ставка аренды (rent_rate)"
        )

    # Конвертируем валюты если необходимо
    property_data_converted = _convert_currencies(property_data, config)

    # Подготавливаем аргументы для calculate_metrics()
    metrics_args = _prepare_metrics_args(property_data_converted, config)

    # Вызываем функцию из cre_analyzer
    metrics = calculate_metrics(**metrics_args)

    return metrics


def _convert_currencies(
    property_data: PropertyData, config: PipelineConfig
) -> PropertyData:
    """
    Конвертирует все денежные значения в базовую валюту (config["currency"]).
    """
    base_currency = config["currency"]
    exchange_rate = config["exchange_rate"]

    # Если все в базовой валюте — ничего не конвертируем
    if (
        property_data.get("value_currency") == base_currency
        and property_data.get("rent_rate_currency") == base_currency
        and property_data.get("opex_currency") == base_currency
    ):
        return property_data

    # Копируем данные
    converted = dict(property_data)

    # Конвертируем value если нужно
    if converted.get("value_currency") != base_currency:
        if base_currency == "USD" and converted.get("value_currency") == "UAH":
            converted["value"] = converted["value"] / exchange_rate
        elif base_currency == "UAH" and converted.get("value_currency") == "USD":
            converted["value"] = converted["value"] * exchange_rate
        converted["value_currency"] = base_currency

    # Конвертируем rent_rate если нужно
    if converted.get("rent_rate") and converted.get("rent_rate_currency") != base_currency:
        if base_currency == "USD" and converted.get("rent_rate_currency") == "UAH":
            converted["rent_rate"] = converted["rent_rate"] / exchange_rate
        elif base_currency == "UAH" and converted.get("rent_rate_currency") == "USD":
            converted["rent_rate"] = converted["rent_rate"] * exchange_rate
        converted["rent_rate_currency"] = base_currency

    # Конвертируем opex если нужно
    if converted.get("opex") and converted.get("opex_currency") != base_currency:
        if base_currency == "USD" and converted.get("opex_currency") == "UAH":
            converted["opex"] = converted["opex"] / exchange_rate
        elif base_currency == "UAH" and converted.get("opex_currency") == "USD":
            converted["opex"] = converted["opex"] * exchange_rate
        converted["opex_currency"] = base_currency

    return converted


def _prepare_metrics_args(
    property_data: PropertyData, config: PipelineConfig
) -> Dict[str, Any]:
    """
    Маппит поля PropertyData в аргументы calculate_metrics().
    """
    args = {
        "property_type": property_data.get("property_type", "інше"),
        "area": property_data.get("gla") or property_data.get("gba"),
        "gba": property_data.get("gba"),
        "gla": property_data.get("gla") or property_data.get("gba"),
        "rent_rate": property_data.get("rent_rate"),
        "vacancy_rate": property_data.get("vacancy_rate", 0.05),
        "opex": property_data.get("opex") or 0,
        "value": property_data.get("value"),
        "rent_growth_rate": property_data.get("rent_growth_rate", 0.03),
        "hold_period": property_data.get("hold_period", 10),
        "exit_cap_rate": property_data.get("exit_cap_rate", 0.09),
        "currency_config": {
            "base": config["currency"],
            "exchange_rate": config["exchange_rate"],
            "inputs": {
                "value": property_data.get("value_currency"),
                "rent_rate": property_data.get("rent_rate_currency"),
                "opex": property_data.get("opex_currency"),
            },
        },
    }

    return args


if __name__ == "__main__":
    # Пример использования
    import json

    test_data = PropertyData(
        property_name="Склад Тест",
        property_type="склад",
        address="вул. Тестова, 1",
        city="Київ",
        description="Тестовий складський комплекс",
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

    config = PipelineConfig(
        object_category="income",
        doc_type="full",
        pres_format="pptx",
        exchange_rate=41.5,
        currency="USD",
    )

    metrics = analyze_property(test_data, config)
    print(json.dumps(metrics, indent=2, ensure_ascii=False, default=str))
