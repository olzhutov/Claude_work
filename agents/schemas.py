"""
Схемы данных для обработки информации об объектах недвижимости.
TypedDict определяют контракт между агентами.
"""

from typing import TypedDict, Optional


class PropertyData(TypedDict):
    """
    Структурированные данные об объекте недвижимости,
    извлечённые из исходного документа (PDF, изображение, текст).
    """

    # === Идентификация объекта ===
    property_name: str
    """Название объекта или адрес (например, "Склад на Борщагівській")"""

    property_type: str
    """Тип объекта: "склад", "офис", "ритейл", "виробництво", "земля", "інше" """

    address: Optional[str]
    """Адрес или геолокация"""

    city: str
    """Город / населённый пункт"""

    description: str
    """Краткое описание объекта из документа"""

    # === Площадь ===
    gba: float
    """Gross Building Area (валовая площадь), м²"""

    gla: Optional[float]
    """Gross Leasable Area (арендуемая площадь), м². Если None → равно gba"""

    # === Финансы (для доходных объектов) ===
    value: float
    """Стоимость объекта (цена предложения)"""

    value_currency: str
    """Валюта стоимости: "USD" или "UAH" """

    rent_rate: Optional[float]
    """Ставка аренды, [currency]/м²/мес"""

    rent_rate_currency: str
    """Валюта аренды: "USD" или "UAH" """

    vacancy_rate: float
    """Коэффициент вакансии 0.0–1.0 (дефолт: 0.05)"""

    opex: Optional[float]
    """Операционные расходы в год"""

    opex_currency: str
    """Валюта расходов: "USD" или "UAH" """

    # === Параметры прогноза (для IRR/DCF анализа) ===
    rent_growth_rate: float
    """Ежегодный рост аренды (дефолт: 0.03)"""

    exit_cap_rate: Optional[float]
    """Cap Rate при продаже объекта (дефолт: 0.09)"""

    hold_period: int
    """Горизонт владения, лет (дефолт: 10)"""

    # === Техническое состояние ===
    year_built: Optional[int]
    """Год постройки"""

    condition: Optional[str]
    """Состояние: "нова", "задовільний", "потребує ремонту" """

    land_area: Optional[float]
    """Площадь земельного участка, га"""

    infrastructure: Optional[str]
    """Описание локации и инфраструктуры"""

    # === Техническая информация (специфично для типа объекта) ===
    property_class: Optional[str]
    """Класс объекта: A, B+, B, C (для офисов/складов)"""

    ceiling_height: Optional[float]
    """Высота потолков, м (для складов/производства)"""

    power_capacity_kva: Optional[float]
    """Электромощность, кВА"""

    loading_docks: Optional[int]
    """Количество погрузочных доков / ворот (для складов)"""

    crane_capacity_tons: Optional[float]
    """Грузоподъёмность кран-балки, т (для производства)"""

    distance_to_highway_km: Optional[float]
    """Расстояние до КАД / основной магистрали, км"""

    # === Земельный участок ===
    land_cadastre: Optional[str]
    """Кадастровый номер участка"""

    land_ownership: Optional[str]
    """Форма собственности участка: "собственность", "аренда" """

    land_lease_years: Optional[int]
    """Срок аренды участка, лет (если аренда)"""

    land_category: Optional[str]
    """Категория земель и целевое назначение"""

    # === Правовой статус ===
    ownership_form: Optional[str]
    """Форма собственности: "частная", "государственная", "смешанная" """

    legal_encumbrances: Optional[str]
    """Наличие обременений, залогов, ограничений (описание)"""

    permits_available: Optional[bool]
    """Есть ли необходимые разрешительные документы"""

    # === Рыночные данные ===
    market_rent_rate: Optional[float]
    """Рыночная ставка аренды для сегмента, [currency]/м²/мес"""

    market_vacancy_rate: Optional[float]
    """Вакансия на рынке сегмента, % (0.0-1.0)"""

    market_trends: Optional[str]
    """Описание трендов на рынке и прогноз"""

    competitors_info: Optional[str]
    """Информация об основных конкурентах (объекты-аналоги)"""

    # === Оценка рисков ===
    risk_legal: Optional[str]
    """Уровень юридического риска: "Высокий", "Средний", "Низкий" + описание"""

    risk_market: Optional[str]
    """Уровень рыночного риска: "Высокий", "Средний", "Низкий" + описание"""

    risk_technical: Optional[str]
    """Уровень технического риска: "Высокий", "Средний", "Низкий" + описание"""

    # === Документы и приложения ===
    documents_list: Optional[list]
    """Список предоставленных документов (PDF, Excel, Word и т.д.)"""

    photos_count: Optional[int]
    """Количество фотографий объекта"""

    plans_available: Optional[bool]
    """Есть ли планировки / чертежи"""

    # === Метаданные извлечения ===
    source_file: str
    """Путь к исходному файлу или папке объекта"""

    extraction_confidence: str
    """Уверенность извлечения: "high" / "medium" / "low" """

    extraction_notes: str
    """Заметки о том, что не удалось извлечь или было предполагаемо"""

    last_updated: Optional[str]
    """Дата последнего обновления данных (ISO format)"""


class PipelineConfig(TypedDict):
    """
    Конфигурация для запуска pipeline.
    Передаётся через CLI аргументы.
    """

    # === Тип объекта ===
    object_category: str
    """
    "income" — доходный объект (с финансовой моделью)
    "prospect" — перспективный объект (без финансовой модели)
    """

    # === Тип документа ===
    doc_type: str
    """
    "teaser" — тизер (3-5 стр)
    "memo" — меморандум (5-10 стр)
    "full" — инвест презентация (10+ стр)
    """

    # === Формат презентации ===
    pres_format: str
    """
    "pptx" — PowerPoint файл
    "gamma" — Gamma (текстовый аутлайн)
    """

    # === Параметры валюты ===
    exchange_rate: float
    """Курс UAH/USD"""

    currency: str
    """Базовая валюта: "USD" или "UAH" """


class BtiRoom(TypedDict):
    """Одне приміщення з експлікації БТІ."""

    id: str
    """Номер приміщення: "1", "2а", "16" """

    name: str
    """Назва: "кімната", "коридор", "санвузол", "сходи" """

    area: float
    """Площа, м²"""

    area_type: str
    """Тип площі: "основна" / "допоміжна" """

    source: str
    """Справа/блок звідки взято: "Справа №1 — МЗК" """


class BtiSprava(TypedDict):
    """Одна Справа (секція) технічного паспорту БТІ."""

    name: str
    """Назва: "Справа №1 — МЗК", "Справа №2 — підвал" """

    total_area: float
    """Загальна площа Справи, м²"""

    usable_area: float
    """Корисна (основна) площа, м²"""

    auxiliary_area: float
    """Допоміжна площа, м²"""

    rooms: list
    """Список приміщень (list[BtiRoom])"""


class BtiAreaData(TypedDict):
    """
    Верифіковані дані про площі об'єкта з документів БТІ.
    Пріоритет: експлікація > ДРРП > план/креслення.
    """

    total_area: float
    """Загальна площа по всіх Справах, м²"""

    usable_area: float
    """Корисна площа, м²"""

    auxiliary_area: float
    """Допоміжна площа, м²"""

    spravy: list
    """Список всіх Справ (list[BtiSprava])"""

    drrp_area: Optional[float]
    """Площа за Витягом ДРРП (якщо є), м²"""

    drrp_match: Optional[bool]
    """True якщо БТІ ≈ ДРРП (±0.5 м²)"""

    discrepancies: list
    """Розбіжності план ≠ експлікація: [{"room": "...", "plan": X, "expl": Y}]"""

    source_files: list
    """Файли, які були прочитані"""

    bti_date: Optional[str]
    """Дата технічного паспорту"""

    confidence: str
    """Впевненість: "high" / "medium" / "low" """

    notes: str
    """Примітки: що знайдено, що відсутнє"""


class AnalysisMetrics(TypedDict):
    """
    Результаты финансового анализа (из cre_analyzer.py).
    Словарь с ключами NOI, Cap Rate, сценариями и т.д.
    (точная структура зависит от cre_analyzer.calculate_metrics())
    """

    pass  # Это просто типизация для dict из cre_analyzer
