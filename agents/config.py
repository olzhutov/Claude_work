"""
Конфигурация для агентов: переменные окружения и константы.
"""

import os

# API ключ Anthropic для работы с Claude API
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

if not ANTHROPIC_API_KEY:
    raise EnvironmentError(
        "Необходимо установить переменную окружения ANTHROPIC_API_KEY. "
        "Получите ключ на console.anthropic.com и выполните: "
        'export ANTHROPIC_API_KEY="sk-ant-..."'
    )

# Модель Claude для извлечения данных
CLAUDE_MODEL = "claude-opus-4-1-20250805"

# Директория для выходных файлов
OUTPUT_DIR = "output"

# Константы валют
SUPPORTED_CURRENCIES = {"USD", "UAH"}
DEFAULT_CURRENCY = "USD"
DEFAULT_EXCHANGE_RATE = 41.5  # UAH/USD

# Notion — интеграция
NOTION_TOKEN = os.environ.get("NOTION_TOKEN")  # Bearer токен из Notion Integrations

# CRE раздел (родительская страница для всех объектов)
NOTION_CRE_SECTION_ID = "32b3dbdd-6ada-81cc-915f-d6f49bd1692f"

# Шаблон страницы объекта (используется как reference для структуры)
NOTION_OBJECT_TEMPLATE = "32b3dbdd-6ada-8121-a854-c109f1858aa6"

# База данных объектов (основная таблица CRE)
NOTION_DATABASE_ID = "32b3dbdd-6ada-8086-840d-d6c5ece812ea"

# ID источника данных (datasource для inline-базы)
NOTION_DATASOURCE_ID = "32b3dbdd-6ada-801d-b380-000b8f0802fb"
