"""
Агент-экстрактор: читает исходные документы и извлекает структурированные данные.
Использует Claude API с vision для обработки PDF, изображений и текстовых файлов.
Поддерживает как единичные файлы, так и папки объектов.
"""

import base64
import json
from pathlib import Path
from typing import Optional, Dict, Any

from anthropic import Anthropic

from config import ANTHROPIC_API_KEY, CLAUDE_MODEL
from schemas import PropertyData


def extract_property_data(file_path: str) -> PropertyData:
    """
    Читает файл и извлекает структурированные данные об объекте.

    Args:
        file_path: Путь к файлу (PDF, изображение, текст)

    Returns:
        PropertyData: Структурированные данные об объекте

    Raises:
        FileNotFoundError: Если файл не найден
        ValueError: Если Claude не смог вернуть валидный JSON
    """
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    # Читаем содержимое файла
    content = _read_file(file_path)

    # Вызываем Claude API для извлечения данных
    client = Anthropic()

    system_prompt = """Ты — эксперт в анализе коммерческой недвижимости.
Тебе предоставляется документ об объекте. Требуется извлечь из него структурированные данные
и вернуть валидный JSON в формате PropertyData.

ПРАВИЛА:
1. Все числовые значения — без кавычек (float/int)
2. Текстовые значения — в кавычках (string)
3. Если поле не найдено — использовать null
4. Для дефолтов применять:
   - vacancy_rate: 0.05
   - rent_growth_rate: 0.03
   - exit_cap_rate: 0.09
   - hold_period: 10
5. extraction_confidence: "high" если уверен, "medium" если частично, "low" если недостаточно данных
6. extraction_notes: описать что не удалось извлечь или что предполагалось

ОБЯЗАТЕЛЬНЫЕ ПОЛЯ (если не найдены — вернуть ошибку):
- property_name, property_type, city, gba, value, value_currency

ВОЗВРАЩАЙ ТОЛЬКО JSON БЕЗ ДОПОЛНИТЕЛЬНОГО ТЕКСТА.
"""

    user_message = f"""Извлеки данные об объекте из документа:

{content}

Вернуть JSON в формате PropertyData с полями:
{{
    "property_name": "...",
    "property_type": "...",
    "address": "...",
    "city": "...",
    "description": "...",
    "gba": 0,
    "gla": null,
    "value": 0,
    "value_currency": "USD",
    "rent_rate": null,
    "rent_rate_currency": "USD",
    "vacancy_rate": 0.05,
    "opex": null,
    "opex_currency": "USD",
    "rent_growth_rate": 0.03,
    "exit_cap_rate": 0.09,
    "hold_period": 10,
    "year_built": null,
    "condition": null,
    "land_area": null,
    "infrastructure": null,
    "source_file": "{file_path}",
    "extraction_confidence": "high",
    "extraction_notes": ""
}}
"""

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=2000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )

    # Парсим JSON ответ
    response_text = response.content[0].text

    # Убираем markdown код-блоки если есть
    if response_text.startswith("```"):
        # Удаляем ```json или ``` в начале
        response_text = response_text.lstrip("`").lstrip("json").lstrip("`").strip()
    if response_text.endswith("```"):
        # Удаляем ``` в конце
        response_text = response_text.rstrip("`").strip()

    try:
        data = json.loads(response_text)
    except json.JSONDecodeError as e:
        raise ValueError(f"Claude вернул некорректный JSON: {response_text}") from e

    # Нормализуем названия полей
    data = _normalize_property_data(data)
    return data


def _read_file(file_path: Path) -> str:
    """
    Читает файл и возвращает содержимое.
    Поддерживает текстовые файлы, PDF и изображения через base64.
    """
    suffix = file_path.suffix.lower()

    # Текстовые файлы
    if suffix in {".txt", ".md"}:
        return file_path.read_text(encoding="utf-8")

    # PDF файлы
    elif suffix == ".pdf":
        return _read_pdf_as_text(file_path)

    # Изображения (через base64)
    elif suffix in {".jpg", ".jpeg", ".png", ".gif", ".webp"}:
        return _read_image_as_base64(file_path)

    else:
        raise ValueError(f"Неподдерживаемый формат файла: {suffix}")


def _read_pdf_as_text(file_path: Path) -> str:
    """
    Читает PDF файл.
    Требует PyPDF2 для полной функциональности; в MVP возвращаем плейсхолдер.
    """
    try:
        import PyPDF2

        text_content = []
        with open(file_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text_content.append(page.extract_text())
        return "\n".join(text_content)
    except ImportError:
        # Если PyPDF2 не установлена, читаем как бинарный и возвращаем ошибку
        raise ImportError(
            "Для обработки PDF требуется установить: pip3 install PyPDF2"
        )


def _normalize_property_data(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Нормализует названия полей из различных форматов Claude в стандартные PropertyData.
    Поддерживает гибкость при разных форматах ответа.
    """
    # Маппинг альтернативных названий полей
    field_mapping = {
        # Площадь
        "construction_year": "year_built",
        "built_year": "year_built",
        "year_constructed": "year_built",

        # Аренда
        "base_rent_monthly": "rent_rate",
        "monthly_rent": "rent_rate",
        "monthly_rent_rate": "rent_rate",
        "asking_rent_per_sqm": "rent_rate",
        "rental_rate": "rent_rate",
        "monthly_rental_rate": "rent_rate",

        # Операционные расходы
        "operational_expenses_yearly": "opex",
        "operating_expenses": "opex",
        "annual_operating_expenses": "opex",
        "yearly_operating_expenses": "opex",
        "operational_expenses": "opex",

        # Вакансия (конвертируем из процентов в доли)
        "vacancy_rate_percent": "vacancy_rate",

        # Рост аренды
        "rental_growth_rate": "rent_growth_rate",
        "annual_rental_growth": "rent_growth_rate",

        # Cap Rate при выходе
        "exit_cap_rate_percent": "exit_cap_rate",
    }

    # Применяем маппинг
    normalized = dict(data)
    for old_key, new_key in field_mapping.items():
        if old_key in normalized and new_key not in normalized:
            value = normalized.pop(old_key)

            # Специальная обработка для процентов
            if old_key in ["vacancy_rate_percent"] and isinstance(value, (int, float)):
                # Конвертируем процент в долю (5 -> 0.05)
                value = value / 100 if value > 1 else value

            normalized[new_key] = value

    # Убедимся что rent_rate в формате $/м²/мес
    if "rent_rate" in normalized and isinstance(normalized["rent_rate"], (int, float)):
        rent = normalized["rent_rate"]
        # Если rent выглядит как годовой (27000), конвертируем в месячный ($ за м²/мес)
        if rent > 100 and "gla" in normalized:
            gla = normalized.get("gla", normalized.get("gba", 1))
            if gla > 0:
                normalized["rent_rate"] = rent / gla / 12

    # Убедимся что vacancy_rate в формате доли (0.05), а не процент (5)
    if "vacancy_rate" in normalized and isinstance(normalized["vacancy_rate"], (int, float)):
        vacancy = normalized["vacancy_rate"]
        # Если > 1, то это процент - конвертируем в долю
        if vacancy > 1:
            normalized["vacancy_rate"] = vacancy / 100

    # Убедимся что exit_cap_rate в формате доли (0.09), а не процент (9)
    if "exit_cap_rate" in normalized and isinstance(normalized["exit_cap_rate"], (int, float)):
        cap_rate = normalized["exit_cap_rate"]
        # Если > 1, то это процент - конвертируем в долю
        if cap_rate > 1:
            normalized["exit_cap_rate"] = cap_rate / 100

    # Убедимся что rent_growth_rate в формате доли (0.03), а не процент (3)
    if "rent_growth_rate" in normalized and isinstance(normalized["rent_growth_rate"], (int, float)):
        growth = normalized["rent_growth_rate"]
        # Если > 1, то это процент - конвертируем в долю
        if growth > 1:
            normalized["rent_growth_rate"] = growth / 100

    return normalized


def _read_image_as_base64(file_path: Path) -> str:
    """
    Кодирует изображение в base64 для отправки в Claude API.
    """
    with open(file_path, "rb") as f:
        image_data = base64.b64encode(f.read()).decode("utf-8")

    suffix = file_path.suffix.lower()
    media_type_map = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".gif": "image/gif",
        ".webp": "image/webp",
    }
    media_type = media_type_map.get(suffix, "image/jpeg")

    # Возвращаем в формате, который Claude API может использовать
    return f"[Изображение в формате {media_type}]\nData: {image_data[:100]}..."


def extract_from_folder(folder_path: str) -> PropertyData:
    """
    Извлекает данные из всех файлов в папке объекта (структура raw/).
    Объединяет информацию из нескольких документов в один PropertyData.

    Args:
        folder_path: Путь к папке объекта (должна содержать raw/ директорию)

    Returns:
        PropertyData: Объединённые структурированные данные

    Raises:
        FileNotFoundError: Если folder не содержит raw/ директорию
    """
    from folder_manager import list_raw_files, get_object_id_from_path

    folder = Path(folder_path)
    raw_dir = folder / "raw"

    if not raw_dir.exists():
        raise FileNotFoundError(f"Папка {raw_dir} не найдена")

    # Получаем все файлы из raw/
    raw_files = list_raw_files(str(folder))

    if not raw_files:
        raise FileNotFoundError(f"Файлы не найдены в {raw_dir}")

    # Собираем содержимое всех файлов
    combined_content = []
    for file_path in raw_files:
        rel_path = file_path.relative_to(raw_dir)
        combined_content.append(f"=== {rel_path} ===\n")

        try:
            if file_path.suffix.lower() in {".txt", ".md"}:
                content = file_path.read_text(encoding="utf-8")
                combined_content.append(content)
            elif file_path.suffix.lower() == ".pdf":
                # Для PDF пытаемся извлечь текст
                content = _read_pdf_as_text(file_path)
                combined_content.append(content)
            elif file_path.suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".webp"}:
                # Для изображений указываем что это изображение
                combined_content.append(f"[Изображение: {file_path.name}]")
        except Exception as e:
            combined_content.append(f"[Ошибка при чтении: {str(e)}]")

        combined_content.append("")

    # Объединённый контент
    full_content = "\n".join(combined_content)

    # Извлекаем данные из объединённого контента
    client = Anthropic()

    system_prompt = """Ты — эксперт в анализе коммерческой недвижимости.
Тебе предоставляются НЕСКОЛЬКО документов об одном объекте недвижимости.
Требуется объединить информацию из всех документов и вернуть структурированные данные.

ПРАВИЛА:
1. Объедини информацию из всех документов логически
2. Приоритет: более свежие данные перевешивают старые
3. Все числовые значения — без кавычек (float/int)
4. Если информация конфликтует — выбери более надежный источник
5. extraction_confidence: "high" если много данных, "medium" если частично, "low" если недостаточно
6. extraction_notes: описать какие документы использованы и что не удалось найти

ОБЯЗАТЕЛЬНЫЕ ПОЛЯ:
- property_name, property_type, city, gba, value, value_currency

ВОЗВРАЩАЙ ТОЛЬКО JSON БЕЗ ДОПОЛНИТЕЛЬНОГО ТЕКСТА.
"""

    user_message = f"""Извлеки данные об объекте из комплекта документов:

{full_content}

Вернуть полный JSON в формате PropertyData с учётом ВСЕХ доступных полей.
"""

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=3000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )

    response_text = response.content[0].text

    # Убираем markdown код-блоки если есть
    if response_text.startswith("```"):
        response_text = response_text.lstrip("`").lstrip("json").lstrip("`").strip()
    if response_text.endswith("```"):
        response_text = response_text.rstrip("`").strip()

    try:
        data = json.loads(response_text)
    except json.JSONDecodeError as e:
        raise ValueError(f"Claude вернул некорректный JSON: {response_text}") from e

    # Нормализуем названия полей
    data = _normalize_property_data(data)

    # Добавляем информацию об источнике (папке)
    object_id = get_object_id_from_path(folder_path)
    if "source_file" not in data:
        data["source_file"] = str(folder_path)

    return data


if __name__ == "__main__":
    # Пример использования
    import sys

    if len(sys.argv) < 2:
        print("Использование: python3 extractor.py <file_path или folder_path>")
        sys.exit(1)

    path = sys.argv[1]
    path_obj = Path(path)

    # Проверяем что это — папка с объектом или файл
    if path_obj.is_dir() and (path_obj / "raw").exists():
        print(f"Обработка папки объекта: {path}")
        data = extract_from_folder(path)
    else:
        print(f"Обработка файла: {path}")
        data = extract_property_data(path)

    print(json.dumps(data, indent=2, ensure_ascii=False))
