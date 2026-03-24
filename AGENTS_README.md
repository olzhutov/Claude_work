# Система агентов для анализа коммерческой недвижимости

## Архитектура системы

### Структура проекта

```
.
├── CLAUDE.md                  # Требования проекта
├── cre_analyzer.py            # Библиотека финансовых расчётов
├── agents/                    # Агенты обработки
│   ├── __init__.py
│   ├── config.py              # Конфигурация (API ключи, константы)
│   ├── schemas.py             # Схемы данных (TypedDict)
│   ├── extractor.py           # Извлечение данных из документов
│   ├── analyzer.py            # Финансовый анализ (использует cre_analyzer)
│   ├── notion_publisher.py    # Генерация Notion-справки (9 разделов)
│   ├── memo_generator.py      # Генерация инвестиционного меморандума
│   ├── presentation_builder.py # Генерация презентации (PPTX или Gamma)
│   ├── folder_manager.py      # Управление структурой папок объектов
│   └── pipeline.py            # Главный оркестратор (CLI)
└── data/
    └── objects/               # Папки объектов недвижимости
        └── {object_id}/
            ├── raw/           # Первичные документы от собственника
            │   ├── documents/ # PDF, DOC, XLSX, TXT
            │   ├── photos/    # JPG, PNG, WEBP
            │   └── plans/     # Планировки, чертежи
            ├── extracted/     # Структурированные данные (JSON)
            └── output/        # Готовые документы (MD, TXT, PPTX)
```

### Два режима работы

**Режим 1 (старый): Один файл**
```bash
python3 agents/pipeline.py test_document.txt --category income --doc-type memo
```
Выходные файлы в `output/`

**Режим 2 (новый, рекомендуется): Папка объекта**
```bash
python3 agents/pipeline.py data/objects/warehouse_kyiv_001 --category income --doc-type memo
```
Выходные файлы в `data/objects/warehouse_kyiv_001/output/`
Извлечённые данные сохраняются в `data/objects/warehouse_kyiv_001/extracted/property_data.json`

## Предварительная настройка

### 1. Получить Claude API ключ
```bash
# Зайти на https://console.anthropic.com
# Settings → API Keys → Create Key
# Копируем ключ
```

### 2. Установить зависимости
```bash
pip3 install anthropic python-pptx
```

### 3. Установить переменную окружения
```bash
# Добавить в ~/.zshrc или ~/.bashrc
export ANTHROPIC_API_KEY="sk-ant-..."

# Применить изменения
source ~/.zshrc
```

### 4. Проверить установку
```bash
cd "/Users/zhutovoleg/Doc/Claude Code"
python3 -c "from agents.config import ANTHROPIC_API_KEY; print('✓ API ключ установлен')"
```

## Быстрый старт

### Шаг 1: Установка API ключа
```bash
# Получить ключ на https://console.anthropic.com
export ANTHROPIC_API_KEY="sk-ant-YOUR_KEY_HERE"
```

### Шаг 2: Тестирование с готовой папкой объекта
```bash
cd "/Users/zhutovoleg/Doc/Claude Code"

# Папка уже создана с тестовым документом
python3 agents/pipeline.py data/objects/warehouse_kyiv_001 \
    --category income \
    --doc-type memo \
    --format gamma \
    --currency USD
```

### Шаг 3: Проверка результатов
```bash
ls -la data/objects/warehouse_kyiv_001/output/
cat data/objects/warehouse_kyiv_001/output/info_brief.md
```

---

## Полный синтаксис команд

### Режим 1: Один файл (обратная совместимость)
```bash
python3 agents/pipeline.py /path/to/document.pdf \
    --category income \           # "income" (доходный) или "prospect" (перспективный)
    --doc-type full \             # "teaser" (3-5 стр), "memo" (5-10 стр), "full" (10+ стр)
    --format pptx \               # "pptx" (PowerPoint) или "gamma" (текстовый аутлайн)
    --exchange-rate 41.5 \        # Курс UAH/USD (дефолт: 41.5)
    --currency USD \              # "USD" или "UAH" (дефолт: USD)
    --output-dir ./output         # Директория для выходных файлов (дефолт: output)
```

### Режим 2: Папка объекта (рекомендуется)
```bash
python3 agents/pipeline.py data/objects/warehouse_kyiv_001 \
    --category income \           # "income" или "prospect"
    --doc-type memo \             # "teaser", "memo" или "full"
    --format gamma \              # "pptx" или "gamma"
    --exchange-rate 41.5 \        # Курс UAH/USD
    --currency USD                # "USD" или "UAH"
```

**Особенности режима 2:**
- Автоматически обнаруживает структуру папки (raw/, extracted/, output/)
- Обрабатывает ВСЕ файлы в `raw/` (documents/, photos/, plans/)
- Сохраняет извлечённые данные в `extracted/property_data.json`
- Выходные файлы создаются в `output/`

## Примеры использования

### Пример 1: Тестирование готовой папки объекта
```bash
python3 agents/pipeline.py data/objects/warehouse_kyiv_001 \
    --category income --doc-type memo --format gamma --currency USD
```
Результаты: `data/objects/warehouse_kyiv_001/output/`

### Пример 2: Полная инвест-презентация (PowerPoint)
```bash
python3 agents/pipeline.py data/objects/warehouse_kyiv_002 \
    --category income --doc-type full --format pptx --currency USD
```

### Пример 3: Перспективный объект (без финанализа), тизер (Gamma)
```bash
python3 agents/pipeline.py data/objects/office_new_001 \
    --category prospect --doc-type teaser --format gamma
```

### Пример 4: С украинской валютой и конкретным курсом
```bash
python3 agents/pipeline.py data/objects/retail_kyiv_001 \
    --category income --doc-type memo --currency UAH --exchange-rate 42.0
```

### Пример 5: Один файл (старый режим)
```bash
python3 agents/pipeline.py test_document.pdf \
    --category income --doc-type memo --format gamma --currency USD
```
Результаты: `output/`

### Создание новой папки объекта вручную
```bash
# Создать структуру
mkdir -p data/objects/warehouse_test_001/{raw/{documents,photos,plans},extracted,output}

# Загрузить документы
cp /path/to/documents/* data/objects/warehouse_test_001/raw/documents/
cp /path/to/photos/* data/objects/warehouse_test_001/raw/photos/

# Запустить pipeline
python3 agents/pipeline.py data/objects/warehouse_test_001 \
    --category income --doc-type full --format pptx
```

## Выходные файлы

После запуска в папке `output/` (или `data/objects/{id}/output/`) создаются файлы:

### 1. `info_brief.md` — Информационная справка (Markdown)

**Структура (9 разделов):**

1. **РЕЗЮМЕ** — ключевые показатели одной строкой
2. **ЗАГАЛЬНА ІНФОРМАЦІЯ** — основные параметры в таблице
3. **ТЕХНІЧНІ ХАРАКТЕРИСТИКИ** — адаптируется под тип объекта:
   - Склад/Логистика: потолки (м), доки, электромощность (кВА)
   - Офис: класс (А/Б/В), парковка, системы
   - Производство: кран-балка (т), газоснабжение
4. **ЗЕМЕЛЬНА ДІЛЯНКА** — участок, кадастр, собственность, срок аренды
5. **СТАН ОБ'ЄКТУ ТА ІНФРАСТРУКТУРА** — состояние, год постройки, транспортность
6. **РИНКОВИЙ КОНТЕКСТ** — рынок, конкуренты, тренды, вакансия
7. **ФІНАНСОВІ ПОКАЗНИКИ** — NOI, Cap Rate, сценарии (только для доходных объектов)
8. **ОЦІНКА РИЗИКІВ** — матрица рисков (юридические, рыночные, технические)
9. **ДОКУМЕНТИ ТА ДОДАТКИ** — список документов, фотографии, планировки

**Использование:**
- Скопируйте содержимое в Claude Code
- Попросите создать страницу в Notion через Notion MCP

### 2. `investment_memo.txt` — Инвестиционный меморандум

Текстовый документ с:
- Резюме объекта
- Финансовыми показателями
- Рисками и возможностями
- Рекомендациями

**Использование:**
- Готовый к печати или отправке документ
- Подходит для презентации инвесторам

### 3. `gamma_outline.txt` или `presentation.pptx` — Презентация

**Формат зависит от параметра `--format`:**

- **gamma** → `gamma_outline.txt` (текстовый аутлайн)
  - Скопируйте в Claude Code
  - Попросите создать Gamma-презентацию

- **pptx** → `presentation.pptx` (PowerPoint)
  - Сразу готов к открытию в PowerPoint
  - Размер зависит от `--doc-type`:
    - teaser (3-5 слайдов)
    - memo (5-10 слайдов)
    - full (10+ слайдов)

### 4. `property_data.json` (только в режиме папки объекта)

Сохраняется в `extracted/property_data.json` при обработке папки объекта.

**Содержит:**
- Все извлечённые данные об объекте
- Используется для последующей обработки
- Можно использовать для интеграции с другими системами

## Типы объектов и документов

### Категории объектов
| Категория | Описание | Что считается |
|-----------|---------|--------------|
| **income** | Доходный объект | NOI, Cap Rate, IRR, сценарии |
| **prospect** | Перспективный объект | Характеристики, локация, бизнес-кейсы |

### Типы документов
| Тип | Страницы | Содержание |
|-----|----------|-----------|
| **teaser** | 3-5 | Краткий обзор + финансовый snapshot |
| **memo** | 5-10 | Характеристики + основные метрики |
| **full** | 10+ | Полный анализ + сценарии + риски |

### Форматы презентаций
| Формат | Использование |
|--------|--------------|
| **pptx** | PowerPoint файл (сразу готов) |
| **gamma** | Текстовый аутлайн для Gamma (вставить в Claude Code) |

## Поддерживаемые форматы и типы объектов

### Форматы документов

**Режим 1 (один файл):**
- `.txt`, `.md` — текстовые файлы
- `.pdf` — PDF (требуется PyPDF2)
- `.jpg`, `.jpeg`, `.png`, `.gif`, `.webp` — изображения

**Режим 2 (папка объекта):**
- Обрабатывает ВСЕ форматы из подпапок:
  - `raw/documents/` — текстовые и PDF файлы
  - `raw/photos/` — изображения объекта
  - `raw/plans/` — планировки и чертежи
- Объединяет данные из всех документов в один PropertyData

### Типы объектов (автоматическое распознавание)

Система поддерживает ВСЕ типы коммерческой недвижимости:
- **Склад / Логистика** — показывает потолки, доки, электромощность
- **Офис** — показывает класс, парковку, системы
- **Производство / Завод** — показывает кран-балку, газоснабжение
- **Ритейл / Торговля** — общие показатели
- **Гостинично-ресторанный комплекс** — специфичные показатели
- **Имущественный комплекс** — перечень зданий
- **Земельные участки** — информация о земле
- **Другие типы** — универсальные показатели

Раздел "Технические характеристики" в справке автоматически адаптируется под тип объекта.

## Требования к исходным данным

Для доходных объектов **обязательны:**
- Название объекта
- Тип объекта
- GBA (валовая площадь)
- Стоимость
- Ставка аренды

**Опциональные поля:**
- GLA, год постройки, состояние, инфраструктура
- Операционные расходы, прирост аренды, Cap Rate выхода

## Валюты

Поддерживаются две валюты:
- **USD** (доллар США)
- **UAH** (гривня)

**Правило курса:**
- Курс UAH/USD **ВСЕГДА** передаётся вручную через `--exchange-rate`
- **Запрещено** хардкодить фиксированный курс

Пример:
```bash
--exchange-rate 41.5  # 1 USD = 41.5 UAH
```

## Вывод данных в Notion и Gamma

### Для Notion
1. Откройте `output/notion_brief.md` в текстовом редакторе
2. Скопируйте содержимое
3. Вставьте в Claude Code
4. Попросите: "Создай эту страницу в Notion"
5. Claude Code использует Notion MCP для создания страницы

### Для Gamma
1. Откройте `output/gamma_outline.txt`
2. Скопируйте содержимое (весь текст)
3. Вставьте в Claude Code
4. Попросите: "Создай Gamma-презентацию по этому аутлайну"
5. Claude Code использует Gamma MCP для создания презентации

## Решение проблем

### "ANTHROPIC_API_KEY не установлен"
```bash
# Проверить:
echo $ANTHROPIC_API_KEY

# Если пусто:
export ANTHROPIC_API_KEY="sk-ant-..."

# Добавить в ~/.zshrc для постоянности
echo 'export ANTHROPIC_API_KEY="sk-ant-..."' >> ~/.zshrc
source ~/.zshrc
```

### "Требуется установить PyPDF2"
```bash
pip3 install PyPDF2
```

### "Требуется установить python-pptx"
```bash
pip3 install python-pptx
```

### Claude API возвращает некорректный JSON
- Проверьте что документ содержит корректные данные об объекте
- Попробуйте текстовый файл вместо PDF
- Обратитесь к документам об объекте для дополнительной информации

## Управление папками объектов

### Создание папки объекта автоматически

```python
from agents.folder_manager import create_object_folder

path = create_object_folder("warehouse_kyiv_002")
# Создаёт: data/objects/warehouse_kyiv_002/{raw/{documents,photos,plans},extracted,output}
```

### Проверка что папка имеет правильную структуру

```python
from agents.folder_manager import is_object_folder

if is_object_folder("data/objects/warehouse_kyiv_001"):
    print("✓ Папка готова к обработке")
```

### Список файлов в raw/

```python
from agents.folder_manager import list_raw_files

files = list_raw_files("data/objects/warehouse_kyiv_001")
files_docs = list_raw_files("data/objects/warehouse_kyiv_001", "documents")
```

### Сохранение и загрузка данных

```python
from agents.folder_manager import save_extracted_data, load_extracted_data

# Сохранить
save_extracted_data("data/objects/warehouse_kyiv_001", property_data)

# Загрузить
data = load_extracted_data("data/objects/warehouse_kyiv_001")
```

---

## Архитектура

### Поток данных (один файл)
```
Файл (PDF/фото/текст)
    ↓
[extractor.py]        ← Claude API извлекает PropertyData
    ↓
[analyzer.py]         ← Финансовый анализ (если доходный объект)
    ↓
    ├── [notion_publisher.py]  → info_brief.md
    ├── [memo_generator.py]    → investment_memo.txt
    └── [presentation_builder.py]
        ├── (pptx)  → presentation.pptx
        └── (gamma) → gamma_outline.txt
```

### Поток данных (папка объекта)
```
raw/
├── documents/*.{txt,pdf,docx}
├── photos/*.{jpg,png,webp}
└── plans/*.{pdf,jpg}
    ↓
[extractor.py extract_from_folder()]  ← Claude API объединяет данные
    ↓
[folder_manager.py save_extracted_data()]
    ↓
extracted/property_data.json
    ↓
[analyzer.py]  ← Финансовый анализ
    ↓
    ├── [notion_publisher.py]  → output/info_brief.md
    ├── [memo_generator.py]    → output/investment_memo.txt
    └── [presentation_builder.py]
        ├── (pptx)  → output/presentation.pptx
        └── (gamma) → output/gamma_outline.txt
```

### Типы данных
- **PropertyData** (TypedDict) — структурированные данные об объекте
- **PipelineConfig** (TypedDict) — конфигурация запуска
- **AnalysisMetrics** (dict) — финансовые метрики из cre_analyzer.py

## Дополнительно

### Интеграция с cre_analyzer.py
Агент `analyzer.py` использует функцию `calculate_metrics()` из основной библиотеки для расчёта:
- NOI (чистый операционный доход)
- Cap Rate (ставка капитализации)
- Сценарии оценки
- Сроки окупаемости
- И других финансовых показателей

### Независимые агенты
Каждый агент можно использовать отдельно:

```python
from agents.extractor import extract_property_data
from agents.analyzer import analyze_property

data = extract_property_data("document.pdf")
metrics = analyze_property(data, config)
```

## Установка зависимостей

### Требуемые пакеты

```bash
pip3 install anthropic
```

### Опциональные пакеты

```bash
# Для обработки PDF
pip3 install PyPDF2

# Для генерации PowerPoint (если используется --format pptx)
pip3 install python-pptx
```

### Проверка установки

```bash
python3 -c "
from agents.folder_manager import create_object_folder
from agents.extractor import extract_from_folder
from agents.analyzer import analyze_property
print('✓ Все агенты импортируются успешно')
"
```

---

## Статус реализации

✅ **Завершено (Phase 2):**

- ✓ `folder_manager.py` — управление папками объектов
- ✓ `schemas.py` — расширенная схема PropertyData (25+ полей)
- ✓ `extractor.py` — извлечение из одного файла и из папки (extract_from_folder)
- ✓ `analyzer.py` — финансовый анализ с интеграцией cre_analyzer
- ✓ `notion_publisher.py` — 9-разделовая информационная справка с адаптацией по типу объекта
- ✓ `memo_generator.py` — инвестиционный меморандум с исправлениями парсинга
- ✓ `presentation_builder.py` — генератор презентаций (PPTX/Gamma) с исправлениями
- ✓ `pipeline.py` — главный оркестратор с поддержкой папок объектов
- ✓ `config.py` — использует claude-opus-4-1-20250805
- ✓ Структура `data/objects/warehouse_kyiv_001/` с тестовым документом

**Готово к тестированию!**

Необходимо установить `ANTHROPIC_API_KEY` и запустить pipeline.

---

## Лицензия

Проект "Анализ коммерческой недвижимости" (Open Source)
