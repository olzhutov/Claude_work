# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## SETUP & COMMANDS

### Installation
```bash
pip3 install anthropic python-pptx
export ANTHROPIC_API_KEY="sk-ant-..."
export NOTION_TOKEN="..."   # Optional, for Notion integration
```

### Run the pipeline
```bash
# Recommended: folder mode
python3 agents/pipeline.py data/objects/{object_id} \
    --category income \       # "income" | "prospect"
    --doc-type memo \         # "teaser" | "memo" | "full"
    --format gamma \          # "pptx" | "gamma"
    --exchange-rate 41.5 \
    --currency USD \
    --location-score          # optional

# Legacy: single file mode
python3 agents/pipeline.py /path/to/document.pdf --category income --doc-type full --format pptx
```

### Verify setup
```bash
python3 verify_setup.py
```

### Open web calculator
Open `web_app/index.html` directly in a browser — no server needed.

### Generate a new Sakura PPTX (reference implementation)
```bash
python3 broadway_sakura.py   # outputs broadway_full.pptx
```

---

## ARCHITECTURE

```
cre_analyzer.py          ← Financial engine (NOI, Cap Rate, IRR). Pure Python stdlib.
agents/
  pipeline.py            ← CLI orchestrator. Entry point for all analysis runs.
  extractor.py           ← Reads raw docs/photos via Claude API (vision) → PropertyData JSON
  analyzer.py            ← PropertyData + cre_analyzer → AnalysisMetrics dict
  notion_publisher.py    ← → info_brief.md  (9-section brief, Ukrainian)
  memo_generator.py      ← → investment_memo.txt
  presentation_builder.py← → presentation.pptx or gamma_outline.txt
  location_scorer.py     ← Location evaluation (6 factors, property-type weights 0-10)
  folder_manager.py      ← Creates/reads object folder structure
  schemas.py             ← TypedDicts: PropertyData (60+ fields), PipelineConfig, AnalysisMetrics
  config.py              ← API keys, model name, currency defaults
web_app/index.html       ← Standalone browser calculator (HTML+CSS+JS, no build step)
broadway_sakura.py       ← CANONICAL reference for Sakura PPTX style (copy from here)
```

### Data flow
```
raw/{documents,photos,plans}/ → extractor.py → extracted/property_data.json
                                                        ↓
                                              analyzer.py → AnalysisMetrics
                                                        ↓
                              notion_publisher / memo_generator / presentation_builder
                                                        ↓
                                              output/{info_brief.md, investment_memo.txt, presentation.*}
```

### Object folder convention
```
data/objects/{object_id}/
  raw/documents/    ← PDF, DOCX, XLSX, TXT from owner
  raw/photos/       ← JPG, PNG, WEBP
  raw/plans/        ← Floor plans
  extracted/        ← property_data.json (auto-generated)
  output/           ← All final documents
```

### Key rules for new code
- **No new dependencies** — stdlib only, except `anthropic` and `python-pptx` where needed
- **All outputs in Ukrainian** — even when logic/comments are in Russian
- **Exchange rate never hardcoded** — always passed via `exchange_rate` parameter; raise error if missing when mixing currencies
- **PPTX style = Sakura** — copy components from `broadway_sakura.py`; slide size must be `10" × 5.625"`, font `Nunito Sans`, bg `#0C1828`
- **Gamma style** — accent `#D5B58A`, card bg `#122941`, body `#FCFCFC`
- **DOCX tables** — always full page width, content-proportional columns → skill `docx-tables-cre`

---

# Проект: Анализ коммерческой недвижимости

## Описание проекта

Разработка универсальных инструментов для анализа коммерческой недвижимости и земельных участков. Проект охватывает все типы объектов коммерческого назначения с приоритетом на складскую и производственную недвижимость, имущественные комплексы и офисы. Фокусируется на расчёте ключевых финансовых метрик и помощи в принятии инвестиционных решений.

---

## ПОВЕДЕНИЕ АГЕНТА-АНАЛИТИКА

### Язык
- **Общение с пользователем**: русский язык
- **Все финальные документы** (справки, меморандумы, презентации, Gamma-аутлайны): **строго украинский язык**
- Переменные и функции в коде: английский (best practice)
- Комментарии в коде, docstring'и, README: русский язык

### Главное правило — автономная работа

#### ❌ ЗАПРЕЩЕНО:
- Спрашивать пользователя "где найти информацию"
- Спрашивать "есть ли у вас данные о..."
- Останавливаться из-за нехватки данных
- Спрашивать "какой файл содержит..."
- Ждать подсказок об источниках

#### ✅ ОБЯЗАТЕЛЬНО:
- Действовать самостоятельно до получения конечного результата
- Сообщать что найдено (или не найдено) — не спрашивать где искать
- При отсутствии данных — использовать обоснованные допущения
- Помечать допущения: `⚠️ Припущення: [...]`

### Порядок поиска информации

Всегда соблюдай эту последовательность:

#### 1️⃣ Документы проекта (в первую очередь)
Самостоятельно проверь ВСЕ файлы в папке объекта. Типичный набор:

```
raw/documents/   — инвестпроект, бизнес-план, технический паспорт,
                   дозвільна документація, договори оренди, фінзвіти
raw/photos/      — фото фасадов, интерьеров, прилегающей территории
raw/plans/       — планировки, чертежи, схемы
extracted/       — уже извлечённые данные (property_data.json)
```

Читай самостоятельно: PDF, DOCX, XLSX, TXT, CSV, JPG — все форматы.

#### 2️⃣ Веб-поиск (если в документах нет)
Ищи автономно:
- **Локация**: Google Maps / OpenStreetMap — расположение, трафик, инфраструктура
- **Рынок аренды**: dom.ria.com, olx.ua, 100realty.ua — ставки, аналоги
- **Аналитика рынка**: открытые отчёты CBRE, Colliers, JLL Ukraine
- **Демография и экономика**: ukrstat.gov.ua, сайты городских советов

#### 3️⃣ Обоснованные допущения (если нигде нет)
Зафиксируй допущение и продолжай работу:
```
⚠️ Припущення: Ставку оренди прийнято на рівні $6/м²/міс на основі
   середньоринкових показників для складів класу B у містах 200-500 тис. мешканців.
```

### Начало работы с новым объектом

При получении нового объекта автоматически выполни:
1. `ls -la data/objects/{object_id}/` — найди все доступные файлы
2. Прочитай ключевые документы из `raw/`
3. Сообщи пользователю на русском: что найдено, что готов сделать
4. Приступай к работе без лишних вопросов

### Типовые сценарии

**Сценарий A: Есть полный пакет документов**
```
1. Читаю все файлы из raw/ → извлекаю данные
2. Дополняю веб-поиском (рынок, конкуренты, локация)
3. Готовлю справку / модель / презентацию на украинском
4. Помечаю что из документов, что из сети
```

**Сценарий B: Документов мало или нет**
```
1. Читаю что есть
2. Активно ищу в интернете
3. Где данных нет — фиксирую допущения с обоснованием ⚠️
4. В конце документа — список данных которые стоит уточнить
```

**Сценарий C: Пользователь задаёт вопрос по объекту**
```
1. Отвечаю сразу на русском — не спрашиваю где найти ответ
2. Указываю источник: документ / веб / допущение
3. Предлагаю углубиться в тему если нужно
```

---

## ОСНОВНЫЕ ПРАВИЛА РАЗРАБОТКИ

### 1. Технологический стек

#### Python-версия (серверная/CLI)
- **Язык**: Чистый Python (Python 3.9+)
- **Принцип**: Никаких лишних внешних библиотек
- **Стандартная библиотека**: Используются только встроенные модули Python
- **Исключения**: `anthropic`, `python-pptx`, `PyPDF2` — только когда явно нужны

#### Web-версия (браузерная)
- **Файл**: `web_app/index.html` — полностью автономный файл (HTML + CSS + JavaScript)
- **Хостинг**: Не требует сервера, открывается через `file://` или публикуется как статика
- **Зависимости**: Внешних JS/CSS-библиотек нет (чистые HTML, CSS, vanilla JavaScript)
- **Функциональность**: Дублирует логику расчётов, вычисляет NOI, Cap Rate, IRR в реальном времени
- **Поддерживаемые типы объектов в web-версии**:
  - **Склад (🏭 Warehouse)** — складские комплексы (дефолт: 5000м² GBA, $6/м²/мес, 5% вакансия, $1.2M)
  - **Офис (🏢 Office)** — офисные центры (дефолт: 8000м² GBA, $12/м²/мес, 8% вакансия, $2M)
  - **Ритейл (🛍️ Retail)** — торговые помещения (дефолт: 6000м² GBA, $15/м²/мес, 10% вакансия, $1.8M)

### 2. Контекст предметной области

#### Охват и масштаб
- **Основной охват**: ВСЕ типы коммерческой недвижимости и земельные участки
- **Поддерживаемые типы объектов**:
  - Складские комплексы (1-класс, 2-класс, 3-класс)
  - Производственные помещения и заводы
  - Имущественные комплексы
  - Офисные здания и бизнес-центры
  - Розничные помещения и торговые центры
  - Гостинично-ресторанные комплексы
  - Логистические центры
  - Земельные участки под коммерческое использование
  - Другие типы коммерческой недвижимости

#### ГЛАВНОЕ ПРАВИЛО: Универсальность
- **Принцип**: Все создаваемые инструменты, скрипты и расчёты должны быть **УНИВЕРСАЛЬНЫМИ** и работать для любого типа недвижимости
- **Реализация**:
  - Инструменты должны быть гибкими и параметризованными
  - Поддержка произвольных входных данных без привязки к типу объекта
  - Расширяемость для добавления новых категорий

### 3. Ключевые метрики (KPI)

Формулы NOI, Cap Rate, IRR → skill `cre-financial-metrics`. Реализация в `cre_analyzer.py`.

### 4. Методология разработки
- **Plan Mode**: Перед написанием любого кода ВСЕГДА предложить план
- **Утверждение**: Дождаться одобрения плана перед реализацией
- **Итерация**: Обсудить подход до начала кодирования

### 5. Региональные настройки

#### Целевая локация
- **Страна**: Украина
- **Рынок**: Коммерческая недвижимость Украины

#### Валюты расчётов
- **Поддерживаемые валюты**: Доллар США (USD, символ $) и Гривна (UAH, символ ₴)
- **Базовая валюта по умолчанию**: USD (рекомендуется для объектов дороже $100 000)

#### Правило курса валют
- Курс UAH/USD **ВСЕГДА** передаётся вручную через параметр `exchange_rate`
- **ЗАПРЕЩЕНО** хардкодить фиксированный курс в коде
- Если курс не передан при смешивании валют — вернуть ошибку
- Пример: `exchange_rate=41.5` означает 1 USD = 41.5 UAH

---

## СТРУКТУРА ПРОЕКТА

```
.
├── CLAUDE.md                       # Этот файл (правила и описание)
├── AGENTS_README.md                # Документация по системе агентов
├── cre_analyzer.py                 # Python-библиотека расчётов (универсальная)
├── web_app/
│   └── index.html                  # Веб-калькулятор (HTML + CSS + JS, автономный)
├── agents/
│   ├── __init__.py
│   ├── config.py                   # Конфигурация (API ключи, константы)
│   ├── schemas.py                  # Схемы данных (TypedDict)
│   ├── extractor.py                # Извлечение данных из документов
│   ├── analyzer.py                 # Финансовый анализ (использует cre_analyzer)
│   ├── notion_publisher.py         # Генерация информационной справки (9 разделов)
│   ├── memo_generator.py           # Генерация инвестиционного меморандума
│   ├── presentation_builder.py     # Генерация презентации (PPTX или Gamma)
│   ├── folder_manager.py           # Управление структурой папок объектов
│   └── pipeline.py                 # Главный оркестратор (CLI)
└── data/
    └── objects/
        └── {object_id}/
            ├── raw/
            │   ├── documents/      # PDF, DOCX, XLSX, TXT от собственника
            │   ├── photos/         # JPG, PNG, WEBP
            │   └── plans/          # Планировки, чертежи
            ├── extracted/          # Структурированные данные (JSON)
            └── output/             # Готовые документы (MD, TXT, PPTX)
```

---

## ВЫХОДНЫЕ ДОКУМЕНТЫ (все на украинском языке)

Детальные инструкции по генерации — в skill `cre-info-and-presentation`.

- `info_brief.md` — інформаційна довідка (9 розділів), `agents/notion_publisher.py`
- `investment_memo.txt` — інвестиційний меморандум, `agents/memo_generator.py`
- `presentation.pptx` — PowerPoint стиль Sakura, `agents/presentation_builder.py`
- `gamma_outline.txt` — Markdown-аутлайн для gamma.app, `agents/presentation_builder.py`
- `property_data.json` — витягнуті дані, зберігається в `extracted/`

---

## ЗАПУСК PIPELINE

```bash
# Режим папки объекта (рекомендуется)
python3 agents/pipeline.py data/objects/{object_id} \
    --category income \      # "income" (доходный) или "prospect" (перспективный)
    --doc-type memo \        # "teaser" (3-5 стр), "memo" (5-10), "full" (10+)
    --format gamma \         # "pptx" или "gamma"
    --currency USD           # "USD" или "UAH"

# Пример
python3 agents/pipeline.py data/objects/warehouse_kyiv_001 \
    --category income --doc-type full --format pptx --currency USD
```

---

## СОГЛАШЕНИЯ О КОДЕ

### Стиль
- PEP 8 compliance
- Максимальная длина строки: 100 символов
- Docstring'и и комментарии — на русском языке

### Пример функции
```python
def calculate_noi(gross_revenue, operating_expenses):
    """
    Расчёт чистого операционного дохода (NOI).

    Args:
        gross_revenue: Валовой доход (в денежных единицах)
        operating_expenses: Операционные расходы

    Returns:
        float: Чистый операционный доход
    """
    # NOI = Валовой доход - Операционные расходы
    return gross_revenue - operating_expenses
```

---

## ПРОЦЕСС РАБОТЫ

1. **Задача получена** → Определяю масштаб, читаю папку объекта
2. **Анализ** → Самостоятельно извлекаю данные из документов и веба
3. **План** → Предлагаю подробный план реализации (Plan Mode)
4. **Обсуждение** → Жду одобрения или уточнений
5. **Реализация** → Пишу код / генерирую документы согласно плану
6. **Тестирование** → Проверяю работоспособность

Все вопросы, уточнения и пожелания обсуждаются на русском языке.

---

## АГЕНТ ОЦІНКИ ЛОКАЦІЇ

Детальний алгоритм, профілі ваг (`LOCATION_PROFILES`), формат виходу → skill `location-scorer-cre`.
Реалізація: `agents/location_scorer.py`, активується через `--location-score` в pipeline.

---

*Версія: 3.3 | Мова спілкування: російська | Мова документів: українська*
