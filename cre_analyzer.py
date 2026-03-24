#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Универсальный анализатор коммерческой недвижимости (расширенная версия).

Модуль предназначен для комплексного финансового анализа объектов коммерческой
недвижимости и земельных участков, включая расчёт:
- Ключевых инвестиционных метрик (NOI, Cap Rate, Cash-on-Cash Return)
- Затрат при покупке и их влияния на доходность
- Сценариев по ставкам капитализации
- Влияния кредитного плеча на доходность
- Прогнозных денежных потоков и окупаемости

Поддерживает любые типы коммерческой недвижимости с полной обратной
совместимостью со старым API.
"""


def _calculate_annuity_payment(loan_amount, annual_rate, term_years):
    """
    Расчёт годового аннуитетного платежа по кредиту.

    Формула: A = P × [r(1+r)^n] / [(1+r)^n - 1]
    где P = сумма кредита, r = годовая ставка, n = количество лет

    Args:
        loan_amount (float): Сумма кредита
        annual_rate (float): Годовая процентная ставка (например, 0.10 = 10%)
        term_years (int): Срок кредита в годах

    Returns:
        float: Годовой платёж по кредиту
    """
    if loan_amount == 0 or annual_rate == 0 or term_years == 0:
        return 0

    r = annual_rate
    n = term_years
    numerator = r * ((1 + r) ** n)
    denominator = ((1 + r) ** n) - 1

    return loan_amount * (numerator / denominator)


def _convert_to_base(amount, from_currency, base_currency, exchange_rate):
    """
    Конвертирует сумму из исходной валюты в базовую.

    Соотношение: 1 USD = exchange_rate UAH

    Args:
        amount (float): Сумма для конвертации
        from_currency (str): Исходная валюта ('USD' или 'UAH')
        base_currency (str): Целевая валюта ('USD' или 'UAH')
        exchange_rate (float): Курс UAH/USD (обязателен при смешивании валют)

    Returns:
        float: Сумма в базовой валюте

    Raises:
        ValueError: Если курс не указан при смешивании валют
    """
    # Если валюта одинакова или сумма нулевая
    if from_currency == base_currency or amount == 0:
        return amount

    # Проверка наличия курса при смешивании валют
    if exchange_rate is None or exchange_rate == 0:
        raise ValueError(
            "exchange_rate обязателен при смешивании валют "
            f"({from_currency} -> {base_currency})"
        )

    # Конвертация USD -> UAH
    if from_currency == "USD" and base_currency == "UAH":
        return amount * exchange_rate

    # Конвертация UAH -> USD
    if from_currency == "UAH" and base_currency == "USD":
        return amount / exchange_rate

    # Если валют не знаем, возвращаем как есть
    return amount


def calculate_metrics(
    property_type,
    vacancy_rate,
    opex,
    value,
    pgi=None,  # Потенциальный валовой доход (опционально, может быть рассчитан)
    area=None,  # Для обратной совместимости
    gba=None,  # Gross Building Area (общая площадь)
    gla=None,  # Gross Leasable Area (арендуемая площадь)
    rent_rate=None,  # Ставка аренды, руб./м²/месяц
    acquisition_costs=None,  # Затраты при покупке (dict)
    opex_detail=None,  # Детализация OPEX (dict)
    debt=None,  # Параметры кредита (dict)
    capex=None,  # CAPEX (dict)
    rent_growth_rate=0.0,  # Индексация аренды, %/год
    hold_period=None,  # Горизонт инвестирования, лет
    exit_cap_rate=None,  # Cap Rate при продаже
    currency_config=None  # Настройки валют (dict)
):
    """
    Расчёт основных инвестиционных метрик объекта недвижимости.

    Поддерживает полный спектр финансовых показателей: от базовых (NOI, Cap Rate)
    до продвинутых (Cash-on-Cash Return, окупаемость, сценарии по cap rate).

    Args:
        property_type (str): Тип объекта ('склад', 'офис', 'габ', 'земля')
        pgi (float): Потенциальный валовой доход, руб./год (или будет рассчитан)
        vacancy_rate (float): Коэффициент потерь (0.0–1.0)
        opex (float): Операционные расходы, руб./год (или будет суммирован из deail)
        value (float): Стоимость покупки, руб. (цена предложения)
        area (float, опц.): Площадь (для обратной совместимости)
        gba (float, опц.): Общая площадь объекта, м² (или соток для земли)
        gla (float, опц.): Арендуемая площадь, м²
        rent_rate (float, опц.): Ставка аренды, руб./м²/месяц
        acquisition_costs (dict, опц.): {
            "pension_fund_pct": 0.01,
            "state_duty_pct": 0.01,
            "agent_commission_pct": 0.02,
            "contract_amount": стоимость по договору (если None = value)
        }
        opex_detail (dict, опц.): {
            "land_tax": ...,
            "property_tax": ...,
            "utilities": ...,
            "maintenance": ...,
            "other": ...
        }
        debt (dict, опц.): {
            "loan_amount": сумма кредита,
            "annual_rate": 0.10 (10%),
            "term_years": 10,
            "amortization": True (аннуитет) или False (только проценты)
        }
        capex (dict, опц.): {
            "renovation": 0,  # единовременные затраты
            "annual_reserve": 0,  # абсолютное значение в год
            "annual_reserve_pct": 0  # % от PGI
        }
        rent_growth_rate (float): Ежегодная индексация аренды (0.03 = 3%)
        hold_period (int, опц.): Горизонт владения, лет
        exit_cap_rate (float, опц.): Cap Rate при продаже (для расчёта exit value)

    Returns:
        dict: Словарь со всеми рассчитанными показателями
    """

    # ===== БЛОК НОРМАЛИЗАЦИИ. Конвертация в базовую валюту =====
    # Извлечение параметров валют из currency_config
    base_currency = "UAH"  # Валюта по умолчанию
    exchange_rate_used = None
    inputs_currencies = {}

    if currency_config is not None:
        base_currency = currency_config.get("base", "UAH")
        exchange_rate_used = currency_config.get("exchange_rate")
        inputs_currencies = currency_config.get("inputs", {})

    # Конвертация входных параметров в базовую валюту
    pgi = _convert_to_base(
        pgi,
        inputs_currencies.get("pgi", base_currency),
        base_currency,
        exchange_rate_used
    )
    rent_rate = _convert_to_base(
        rent_rate,
        inputs_currencies.get("rent_rate", base_currency),
        base_currency,
        exchange_rate_used
    )
    opex = _convert_to_base(
        opex,
        inputs_currencies.get("opex", base_currency),
        base_currency,
        exchange_rate_used
    )
    value = _convert_to_base(
        value,
        inputs_currencies.get("value", base_currency),
        base_currency,
        exchange_rate_used
    )

    # Конвертация параметров в словарях (acquisition_costs, debt, capex)
    if acquisition_costs is not None:
        acq_copy = dict(acquisition_costs)
        if "contract_amount" in acq_copy and acq_copy["contract_amount"] is not None:
            acq_copy["contract_amount"] = _convert_to_base(
                acq_copy["contract_amount"],
                inputs_currencies.get("contract_amount", base_currency),
                base_currency,
                exchange_rate_used
            )
        acquisition_costs = acq_copy

    if debt is not None:
        debt_copy = dict(debt)
        if "loan_amount" in debt_copy:
            debt_copy["loan_amount"] = _convert_to_base(
                debt_copy["loan_amount"],
                inputs_currencies.get("loan_amount", base_currency),
                base_currency,
                exchange_rate_used
            )
        debt = debt_copy

    if capex is not None:
        capex_copy = dict(capex)
        if "renovation" in capex_copy:
            capex_copy["renovation"] = _convert_to_base(
                capex_copy["renovation"],
                inputs_currencies.get("capex_renovation", base_currency),
                base_currency,
                exchange_rate_used
            )
        if "annual_reserve" in capex_copy:
            capex_copy["annual_reserve"] = _convert_to_base(
                capex_copy["annual_reserve"],
                inputs_currencies.get("capex_annual_reserve", base_currency),
                base_currency,
                exchange_rate_used
            )
        capex = capex_copy

    # ===== БЛОК A. Разрешение параметров площади =====
    # Обработка обратной совместимости: area → gba
    if gba is None:
        gba = area
    if gba is None:
        return {"ошибка": "Необходимо указать площадь (area или gba)"}

    if gla is None:
        gla = gba

    # Проверка на нулевые значения
    if gba == 0 or value == 0:
        return {
            "ошибка": "Площадь и стоимость должны быть > 0",
            "property_type": property_type,
            "gba": gba,
            "value": value
        }

    # ===== БЛОК A. Расчёт PGI если задана ставка аренды =====
    if rent_rate is not None:
        # PGI = GLA × ставка × 12 месяцев
        pgi = gla * rent_rate * 12
    elif pgi is None:
        return {"ошибка": "Необходимо указать PGI или rent_rate"}

    # ===== Базовые расчёты: EGI, NOI =====
    egi = pgi * (1 - vacancy_rate)
    vacancy_loss = pgi * vacancy_rate

    # Если задана детализация OPEX, суммируем её
    if opex_detail is not None:
        opex = sum(opex_detail.values())

    noi = egi - opex

    # ===== БЛОК B. Затраты при покупке =====
    total_acquisition_cost = value  # по умолчанию
    transaction_costs = 0

    if acquisition_costs is not None:
        contract_amount = acquisition_costs.get("contract_amount", value)
        pension_fund = acquisition_costs.get("pension_fund_pct", 0)
        state_duty = acquisition_costs.get("state_duty_pct", 0)
        agent_commission = acquisition_costs.get("agent_commission_pct", 0)

        # Пенсионный фонд + госпошлина считаются от суммы по договору
        # Комиссия АН считается от полной стоимости объекта
        pension_state = contract_amount * (pension_fund + state_duty)
        agent_fee = value * agent_commission

        transaction_costs = pension_state + agent_fee
        total_acquisition_cost = value + transaction_costs

    # ===== Расчёт Cap Rate от себестоимости покупки =====
    if total_acquisition_cost > 0:
        cap_rate_on_cost = (noi / total_acquisition_cost) * 100
    else:
        cap_rate_on_cost = 0

    # Cap Rate от цены предложения (для совместимости)
    cap_rate = (noi / value) * 100 if value > 0 else 0

    # ===== БЛОК E. CAPEX (влияние на скорректированный NOI) =====
    annual_capex = 0
    if capex is not None:
        if "annual_reserve" in capex and capex["annual_reserve"] > 0:
            annual_capex = capex["annual_reserve"]
        elif "annual_reserve_pct" in capex and capex["annual_reserve_pct"] > 0:
            annual_capex = pgi * capex["annual_reserve_pct"]

    adjusted_noi = noi - annual_capex

    # ===== БЛОК D. Долговое финансирование =====
    equity = total_acquisition_cost  # по умолчанию, весь объект = собственный капитал
    debt_service = 0
    cash_flow = noi  # по умолчанию, весь NOI
    cocr = 0  # Cash-on-Cash Return

    if debt is not None:
        loan_amount = debt.get("loan_amount", 0)
        annual_rate = debt.get("annual_rate", 0)
        term_years = debt.get("term_years", 1)
        use_amortization = debt.get("amortization", True)

        if use_amortization:
            debt_service = _calculate_annuity_payment(loan_amount, annual_rate, term_years)
        else:
            # Только проценты
            debt_service = loan_amount * annual_rate

        equity = total_acquisition_cost - loan_amount
        cash_flow = noi - debt_service

        if equity > 0:
            cocr = (cash_flow / equity) * 100

    # ===== Расчёт окупаемости =====
    payback_years = total_acquisition_cost / noi if noi > 0 else float('inf')

    # ===== БЛОК F. Горизонт инвестирования и exit value =====
    exit_value = None
    if hold_period is not None and exit_cap_rate is not None and hold_period > 0:
        # Рост доходов за период
        noi_at_exit = noi * ((1 + rent_growth_rate) ** hold_period)
        # Стоимость при выходе
        exit_value = noi_at_exit / exit_cap_rate if exit_cap_rate > 0 else 0

    # ===== Единица измерения удельной стоимости =====
    if property_type.lower() == "земля":
        price_per_unit = value / gba
        unit_name = "руб./сотка"
        price_per_hectare = price_per_unit * 100  # 1 га = 100 соток
    else:
        price_per_unit = value / gba
        unit_name = "руб./м²"
        price_per_hectare = None

    # ===== БЛОК G. Таблица сценариев Cap Rate =====
    # Целевые цены при разных cap rate (для справедливой оценки)
    cap_rate_scenarios = {}
    for target_cap in [0.08, 0.10, 0.12, 0.14]:  # 8%, 10%, 12%, 14%
        target_price = noi / target_cap if target_cap > 0 else 0
        ratio = target_price / value if value > 0 else 0
        cap_rate_scenarios[int(target_cap * 100)] = {
            "price": target_price,
            "ratio": ratio
        }

    # ===== Формирование результирующего словаря =====
    result = {
        # Основные параметры объекта
        "property_type": property_type,
        "gba": gba,
        "gla": gla,
        "rent_rate": rent_rate,

        # Доходы
        "pgi": pgi,
        "vacancy_rate": vacancy_rate,
        "vacancy_loss": vacancy_loss,
        "egi": egi,

        # Расходы
        "opex": opex,
        "opex_detail": opex_detail,
        "annual_capex": annual_capex,

        # NOI
        "noi": noi,
        "adjusted_noi": adjusted_noi,

        # Стоимость
        "value": value,
        "transaction_costs": transaction_costs,
        "total_acquisition_cost": total_acquisition_cost,

        # Cap Rate
        "cap_rate": cap_rate,  # от цены предложения
        "cap_rate_on_cost": cap_rate_on_cost,  # от себестоимости

        # Долг
        "debt": debt,
        "equity": equity,
        "debt_service": debt_service,
        "cash_flow": cash_flow,
        "cocr": cocr,

        # Удельная стоимость
        "price_per_unit": price_per_unit,
        "unit_name": unit_name,
        "price_per_hectare": price_per_hectare,

        # Показатели окупаемости
        "payback_years": payback_years,

        # Долгосрочный анализ
        "rent_growth_rate": rent_growth_rate,
        "hold_period": hold_period,
        "exit_cap_rate": exit_cap_rate,
        "exit_value": exit_value,

        # Сценарии
        "cap_rate_scenarios": cap_rate_scenarios,

        # Информация о валютах (БЛОК K)
        "base_currency": base_currency,
        "exchange_rate": exchange_rate_used
    }

    return result


def convert_report(metrics, output_currency, exchange_rate):
    """
    Конвертирует все денежные поля в словаре metrics в указанную валюту.

    Поля-проценты (cap_rate, vacancy_rate, cocr) остаются неизменными.
    Возвращает новый dict с конвертированными значениями.

    Args:
        metrics (dict): Результат функции calculate_metrics()
        output_currency (str): Целевая валюта ('USD' или 'UAH')
        exchange_rate (float): Курс UAH/USD (обязателен если валюты разные)

    Returns:
        dict: Копия metrics с конвертированными денежными полями
    """
    base_currency = metrics.get("base_currency", "UAH")

    # Если валюты одинаковые, возвращаем оригинальные metrics
    if base_currency == output_currency:
        return metrics

    # Список денежных полей (в тысячах денежных единиц)
    money_fields = [
        "pgi", "vacancy_loss", "egi", "opex", "annual_capex",
        "noi", "adjusted_noi", "value", "transaction_costs",
        "total_acquisition_cost", "debt_service", "cash_flow",
        "price_per_unit", "price_per_hectare", "exit_value"
    ]

    # Создаём копию metrics
    converted = dict(metrics)

    # Конвертируем денежные поля
    for field in money_fields:
        if field in converted and converted[field] is not None:
            converted[field] = _convert_to_base(
                converted[field],
                base_currency,
                output_currency,
                exchange_rate
            )

    # Конвертируем словари с денежными значениями
    if "opex_detail" in converted and converted["opex_detail"] is not None:
        opex_detail_conv = {}
        for key, value in converted["opex_detail"].items():
            opex_detail_conv[key] = _convert_to_base(
                value, base_currency, output_currency, exchange_rate
            )
        converted["opex_detail"] = opex_detail_conv

    # Конвертируем сценарии cap_rate
    if "cap_rate_scenarios" in converted and converted["cap_rate_scenarios"]:
        scenarios_conv = {}
        for rate_key, scenario_data in converted["cap_rate_scenarios"].items():
            scenarios_conv[rate_key] = {
                "price": _convert_to_base(
                    scenario_data["price"], base_currency,
                    output_currency, exchange_rate
                ),
                "ratio": scenario_data["ratio"]  # процент не меняется
            }
        converted["cap_rate_scenarios"] = scenarios_conv

    # Обновляем базовую валюту и курс в результате
    converted["base_currency"] = output_currency
    converted["exchange_rate"] = exchange_rate

    return converted


def print_report(metrics, output_currency=None):
    """
    Форматированный вывод полного отчёта об анализе недвижимости.

    Выводит все рассчитанные показатели структурированно с разделителями.

    Args:
        metrics (dict): Результат функции calculate_metrics()
        output_currency (str, опц.): Валюта вывода ('USD' или 'UAH').
            Если не указана → использует базовую валюту из metrics.
    """

    # Проверка на ошибки
    if "ошибка" in metrics:
        print(f"❌ Ошибка: {metrics['ошибка']}")
        return

    # Конвертация в требуемую валюту если указана
    base_currency = metrics.get("base_currency", "UAH")
    if output_currency is not None and output_currency != base_currency:
        exchange_rate = metrics.get("exchange_rate")
        if exchange_rate is None:
            print("⚠️  Внимание: не указан курс для конвертации валют")
            return
        metrics = convert_report(metrics, output_currency, exchange_rate)

    # Определяем символ валюты для вывода
    currency_symbol = "$" if metrics.get("base_currency") == "USD" else "₴"

    # Названия объектов
    property_name_map = {
        "склад": "Складской комплекс",
        "офис": "Офисное здание",
        "габ": "Гостинично-административный комплекс",
        "земля": "Земельный участок"
    }
    property_name = property_name_map.get(
        metrics["property_type"].lower(),
        metrics["property_type"]
    )

    # ===== ЗАГОЛОВОК =====
    print("=" * 80)
    print(f"  АНАЛИЗ ОБЪЕКТА: {property_name}")
    print("=" * 80)

    # ===== БЛОК 1: ОБЪЕКТ И ПЛОЩАДЬ =====
    print("\n  📍 ОПИСАНИЕ ОБЪЕКТА")
    print("-" * 80)
    print(f"  Тип объекта:                   {metrics['property_type']}")

    if metrics["unit_name"] == "руб./м²":
        print(f"  Общая площадь (GBA):           {metrics['gba']:,.0f} м²")
        if metrics['gla'] != metrics['gba']:
            print(f"  Арендуемая площадь (GLA):      {metrics['gla']:,.0f} м²")
    else:
        print(f"  Площадь (GBA):                 {metrics['gba']:,.1f} соток")
        if metrics['gla'] != metrics['gba']:
            print(f"  Арендуемая площадь (GLA):      {metrics['gla']:,.1f} соток")

    if metrics["rent_rate"] is not None:
        print(
            f"  Ставка аренды:                 {metrics['rent_rate']:,.2f} "
            f"{currency_symbol}/м²/месяц"
        )

    # ===== БЛОК 2: ДОХОДЫ =====
    print("\n  💰 ДОХОДЫ")
    print("-" * 80)
    print(
        f"  PGI (потенциальный валовой):   {metrics['pgi']:,.0f} "
        f"{currency_symbol}/год"
    )
    print(
        f"  Вакансия ({metrics['vacancy_rate']*100:.1f}%):              "
        f"{metrics['vacancy_loss']:,.0f} {currency_symbol}/год"
    )
    print(
        f"  EGI (эффективный валовой):     {metrics['egi']:,.0f} "
        f"{currency_symbol}/год"
    )

    # ===== БЛОК 3: РАСХОДЫ =====
    print("\n  📊 РАСХОДЫ")
    print("-" * 80)

    if metrics["opex_detail"] is not None:
        for key, value in metrics["opex_detail"].items():
            label = {
                "land_tax": "Налог на землю",
                "property_tax": "Налог на недвижимость",
                "utilities": "Коммунальные платежи",
                "maintenance": "Обслуживание и ремонт",
                "other": "Прочие расходы"
            }.get(key, key)
            print(
                f"  {label:.<40} {value:>10,.0f} {currency_symbol}/год"
            )

    print(
        f"  OPEX (ИТОГО):                  {metrics['opex']:,.0f} "
        f"{currency_symbol}/год"
    )

    if metrics["annual_capex"] > 0:
        print(
            f"  CAPEX (резерв на капремонт):   {metrics['annual_capex']:,.0f} "
            f"{currency_symbol}/год"
        )

    # ===== БЛОК 4: ДОХОДНОСТЬ =====
    print("\n  📈 ДОХОДНОСТЬ (ДО ФИНАНСИРОВАНИЯ)")
    print("-" * 80)
    print(
        f"  NOI (чистый операционный):     {metrics['noi']:,.0f} "
        f"{currency_symbol}/год"
    )

    if metrics["adjusted_noi"] != metrics["noi"]:
        print(
            f"  NOI (скорректированный):       {metrics['adjusted_noi']:,.0f} "
            f"{currency_symbol}/год"
        )

    print(f"  Cap Rate (от цены):            {metrics['cap_rate']:,.2f}%")
    print(
        f"  Cap Rate (от себестоимости):   {metrics['cap_rate_on_cost']:,.2f}%"
    )
    print(f"  Окупаемость, лет:              {metrics['payback_years']:.1f}")

    # ===== БЛОК 5: ЗАТРАТЫ ПРИ ПОКУПКЕ =====
    if metrics["transaction_costs"] > 0:
        print("\n  🏦 ЗАТРАТЫ ПРИ ПОКУПКЕ")
        print("-" * 80)
        print(f"  Цена предложения:              {metrics['value']:,.0f} {currency_symbol}")
        print(
            f"  Транзакционные расходы:       {metrics['transaction_costs']:,.0f} "
            f"{currency_symbol}"
        )
        print(
            f"  Итого себестоимость:          {metrics['total_acquisition_cost']:,.0f} "
            f"{currency_symbol}"
        )

    # ===== БЛОК 6: ФИНАНСИРОВАНИЕ =====
    if metrics["debt"] is not None:
        print("\n  💳 ФИНАНСИРОВАНИЕ (КРЕДИТ)")
        print("-" * 80)
        loan_amount = metrics["debt"].get("loan_amount", 0)
        print(f"  Сумма кредита:                 {loan_amount:,.0f} {currency_symbol}")
        print(f"  Собственный капитал (Equity):  {metrics['equity']:,.0f} {currency_symbol}")
        print(f"  Годовой платёж по кредиту:    {metrics['debt_service']:,.0f} {currency_symbol}")
        print(
            f"  Денежный поток (NOI - DS):     {metrics['cash_flow']:,.0f} "
            f"{currency_symbol}/год"
        )
        print(f"  Cash-on-Cash Return:           {metrics['cocr']:,.2f}%")

    # ===== БЛОК 7: УДЕЛЬНАЯ СТОИМОСТЬ =====
    print("\n  🏷️  УДЕЛЬНАЯ СТОИМОСТЬ")
    print("-" * 80)
    unit_label = metrics["unit_name"].split("/")[1]
    print(
        f"  Цена за {unit_label}:            {metrics['price_per_unit']:,.0f} "
        f"{currency_symbol}/{unit_label}"
    )

    if metrics["price_per_hectare"] is not None:
        print(
            f"  Цена за гектар:                {metrics['price_per_hectare']:,.0f} "
            f"{currency_symbol}/га"
        )

    # ===== БЛОК 8: ДОЛГОСРОЧНЫЙ АНАЛИЗ =====
    if metrics["hold_period"] is not None and metrics["exit_value"] is not None:
        print("\n  🔮 ДОЛГОСРОЧНЫЙ ПРОГНОЗ")
        print("-" * 80)
        print(f"  Горизонт инвестирования:      {metrics['hold_period']} лет")
        print(f"  Ежегодная индексация:         {metrics['rent_growth_rate']*100:.1f}%")
        print(f"  Cap Rate при выходе:          {metrics['exit_cap_rate']*100:.1f}%")
        print(
            f"  Предполагаемая стоимость:     {metrics['exit_value']:,.0f} "
            f"{currency_symbol}"
        )

    # ===== БЛОК 9: ТАБЛИЦА СЦЕНАРИЕВ =====
    print("\n  📊 ТАБЛИЦА ЦЕЛЕВЫХ ЦЕН (справедливая оценка при разных Cap Rate)")
    print("-" * 80)
    print(
        f"  Cap Rate | Целевая цена, {currency_symbol} | Отношение к предложению | Оценка"
    )
    print(f"  {'-'*7}|{'-'*19}|{'-'*23}|{'-'*10}")

    for cap_pct, data in sorted(metrics["cap_rate_scenarios"].items()):
        ratio = data["ratio"]
        price = data["price"]

        if ratio > 1.1:
            assessment = "▲ Выше"
        elif ratio < 0.9:
            assessment = "▼ Ниже"
        else:
            assessment = "= Равно"

        print(f"  {cap_pct:>6}% | {price:>17,.0f} | {ratio:>21.2f}x | {assessment}")

    # ===== НИЖНИЙ РАЗДЕЛИТЕЛЬ =====
    print("=" * 80)
    print()


if __name__ == "__main__":
    """
    Примеры использования расширенного анализатора.
    """

    # ========== ПРИМЕР 1: БАЗОВЫЙ АНАЛИЗ (обратная совместимость) ==========
    print("\n")
    print("█" * 80)
    print("  ПРИМЕР 1: СКЛАДСКОЙ КОМПЛЕКС (базовый анализ, обратная совместимость)")
    print("█" * 80)
    print("\n")

    warehouse_basic = calculate_metrics(
        property_type="склад",
        vacancy_rate=0.05,
        opex=2_500_000,
        value=100_000_000,
        pgi=12_000_000,
        area=5_000
    )

    print_report(warehouse_basic)

    # ========== ПРИМЕР 2: ЗЕМЕЛЬНЫЙ УЧАСТОК (базовый) ==========
    print("\n")
    print("█" * 80)
    print("  ПРИМЕР 2: ЗЕМЕЛЬНЫЙ УЧАСТОК (базовый анализ)")
    print("█" * 80)
    print("\n")

    land_basic = calculate_metrics(
        property_type="земля",
        vacancy_rate=0.10,
        opex=400_000,
        value=25_000_000,
        pgi=3_600_000,
        area=50
    )

    print_report(land_basic)

    # ========== ПРИМЕР 3: СКЛАДСКОЙ КОМПЛЕКС С ЗАТРАТАМИ И КРЕДИТОМ ==========
    print("\n")
    print("█" * 80)
    print("  ПРИМЕР 3: СКЛАД С ЗАТРАТАМИ ПОКУПКИ И КРЕДИТНЫМ ФИНАНСИРОВАНИЕМ")
    print("█" * 80)
    print("\n")

    warehouse_advanced = calculate_metrics(
        property_type="склад",
        vacancy_rate=0.05,
        opex=2_500_000,
        value=100_000_000,
        pgi=12_000_000,
        gba=5_000,
        gla=5_000,
        acquisition_costs={
            "pension_fund_pct": 0.01,  # 1%
            "state_duty_pct": 0.01,    # 1%
            "agent_commission_pct": 0.02,  # 2%
            "contract_amount": 80_000_000  # занижена сумма по договору
        },
        opex_detail={
            "land_tax": 200_000,
            "property_tax": 400_000,
            "utilities": 300_000,
            "maintenance": 800_000,
            "other": 100_000
        },
        debt={
            "loan_amount": 50_000_000,
            "annual_rate": 0.10,  # 10% годовых
            "term_years": 10,
            "amortization": True
        },
        capex={
            "annual_reserve": 100_000  # резерв на капремонт
        }
    )

    print_report(warehouse_advanced)

    # ========== ПРИМЕР 4: ИМУЩЕСТВЕННЫЙ КОМПЛЕКС С ИНДЕКСАЦИЕЙ ==========
    print("\n")
    print("█" * 80)
    print("  ПРИМЕР 4: ИМУЩЕСТВЕННЫЙ КОМПЛЕКС С ПРОГНОЗОМ И ИНДЕКСАЦИЕЙ")
    print("█" * 80)
    print("\n")

    complex_with_forecast = calculate_metrics(
        property_type="габ",
        rent_rate=35.0,  # 35 руб./м²/месяц (PGI рассчитается как 14000*35*12 = 5,880,000)
        gba=15_000,
        gla=14_000,
        vacancy_rate=0.08,  # 8% вакансия
        opex=2_800_000,  # Операционные расходы
        value=150_000_000,
        acquisition_costs={
            "pension_fund_pct": 0.01,
            "state_duty_pct": 0.01,
            "agent_commission_pct": 0.02
        },
        debt={
            "loan_amount": 75_000_000,
            "annual_rate": 0.095,
            "term_years": 15,
            "amortization": True
        },
        capex={
            "annual_reserve_pct": 0.02  # 2% от PGI на капремонт
        },
        rent_growth_rate=0.03,  # Ежегодный рост аренды на 3%
        hold_period=5,  # Горизонт 5 лет
        exit_cap_rate=0.10  # Выход при 10% cap rate
    )

    print_report(complex_with_forecast)

    # ========== ПРИМЕР 5: МУЛЬТИВАЛЮТНЫЙ АНАЛИЗ (USD + UAH) ==========
    print("\n")
    print("█" * 80)
    print(
        "  ПРИМЕР 5: СКЛАДСКОЙ КОМПЛЕКС "
        "(аренда в USD, OPEX/CAPEX в UAH)"
    )
    print("█" * 80)
    print("\n")

    currency_config = {
        "base": "USD",          # Базовая валюта всех расчётов и вывода
        "exchange_rate": 41.5,  # 1 USD = 41.5 UAH (курс UAH/USD)
        "inputs": {
            "rent_rate": "USD",
            "value": "USD",
            "opex": "UAH",
            "capex_renovation": "UAH"
        }
    }

    warehouse_multicurrency = calculate_metrics(
        property_type="склад",
        vacancy_rate=0.05,
        gba=5_000,
        gla=4_500,
        rent_rate=6.0,          # $6/м²/мес (в долларах)
        opex=900_000,           # ₴900 000/год (в гривнах)
        value=1_200_000,        # $1 200 000 (в долларах)
        capex={"renovation": 1_640_000},  # ₴1 640 000 (в гривнах)
        currency_config=currency_config
    )

    print("  ◀━ Отчёт в базовой валюте (USD) ━▶")
    print_report(warehouse_multicurrency)

    print("\n" * 2)
    print("  ◀━ Тот же отчёт в гривнах (UAH) ━▶")
    print_report(warehouse_multicurrency, output_currency="UAH")
