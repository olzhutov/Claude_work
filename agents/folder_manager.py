"""
Менеджер структуры папок объекта.
Работает с архитектурой Объекты/{Имя}/ — сырые файлы в корне, аналитика в wiki/.
"""

import json
from pathlib import Path
from typing import List, Dict, Any, Optional

# Расширения файлов, которые считаются сырыми документами
RAW_EXTENSIONS = {
    ".pdf", ".docx", ".xlsx", ".xls", ".doc",
    ".txt", ".csv", ".md",
    ".jpg", ".jpeg", ".png", ".webp", ".gif",
}

# Папки внутри объекта, которые НЕ являются сырыми файлами
EXCLUDED_DIRS = {"wiki", ".obsidian", ".DS_Store"}


def is_object_folder(folder_path: str) -> bool:
    """
    Проверяет, что папка является объектом в архитектуре Объекты/.
    Критерий: папка существует и содержит wiki/ подпапку.
    """
    folder = Path(folder_path)
    return folder.is_dir() and (folder / "wiki").is_dir()


def list_raw_files(folder_path: str, subfolder: Optional[str] = None) -> List[Path]:
    """
    Возвращает список сырых файлов из корня папки объекта.
    Игнорирует папку wiki/ и служебные директории.

    Args:
        folder_path: Путь к папке объекта (Объекты/{Имя}/)
        subfolder: Игнорируется в новой архитектуре (для совместимости)

    Returns:
        Список Path к файлам
    """
    folder = Path(folder_path)

    if not folder.exists():
        return []

    files = []
    for file_path in folder.iterdir():
        # Пропускаем директории и служебные файлы
        if file_path.is_dir():
            continue
        if file_path.name.startswith("."):
            continue
        if file_path.suffix.lower() in RAW_EXTENSIONS:
            files.append(file_path)

    return sorted(files)


def get_extracted_path(folder_path: str) -> Path:
    """Возвращает путь к property_data.json (в wiki/financials/)."""
    folder = Path(folder_path)
    return folder / "wiki" / "financials" / "property_data.json"


def save_extracted_data(folder_path: str, data: Dict[str, Any]) -> Path:
    """
    Сохраняет извлечённые данные в wiki/financials/property_data.json.

    Returns:
        Path к сохранённому файлу
    """
    output_path = get_extracted_path(folder_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    return output_path


def load_extracted_data(folder_path: str) -> Optional[Dict[str, Any]]:
    """
    Загружает извлечённые данные из wiki/financials/property_data.json.

    Returns:
        Dict с данными или None если файл не найден
    """
    json_path = get_extracted_path(folder_path)

    if not json_path.exists():
        return None

    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)


def get_output_path(folder_path: str, filename: str) -> Path:
    """
    Возвращает путь для выходного файла в wiki/.

    Маппинг по расширениям:
        .md   → wiki/objects/{filename}
        .txt  → wiki/financials/{filename}
        .pptx → wiki/{filename}
        rest  → wiki/{filename}
    """
    folder = Path(folder_path)
    suffix = Path(filename).suffix.lower()

    if suffix == ".md":
        return folder / "wiki" / "objects" / filename
    elif suffix == ".txt":
        return folder / "wiki" / "financials" / filename
    else:
        return folder / "wiki" / filename


def save_output_file(folder_path: str, filename: str, content: str) -> Path:
    """
    Сохраняет выходной текстовый файл.

    Returns:
        Path к сохранённому файлу
    """
    output_path = get_output_path(folder_path, filename)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)

    return output_path


def get_object_id_from_path(folder_path: str) -> str:
    """
    Извлекает имя объекта из пути.

    Примеры:
        "Объекты/Владимирская_8" → "Владимирская_8"
        "/full/path/Объекты/Фастов_Брандта50" → "Фастов_Брандта50"
    """
    folder = Path(folder_path)
    return folder.name


def create_object_folder(object_name: str, base_path: str = "Объекты") -> Path:
    """
    Создаёт структуру папок для нового объекта в архитектуре Объекты/.

    Args:
        object_name: Имя объекта (например, "Владимирская_8")
        base_path: Базовая директория

    Returns:
        Path к созданной папке объекта

    Структура:
        Объекты/
        └── object_name/
            └── wiki/
                ├── objects/
                ├── tenants/
                ├── contracts/
                ├── financials/
                └── topics/
    """
    object_dir = Path(base_path) / object_name

    for subdir in ["objects", "tenants", "contracts", "financials", "topics"]:
        (object_dir / "wiki" / subdir).mkdir(parents=True, exist_ok=True)

    return object_dir


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        path = sys.argv[1]
        print(f"Объект: {get_object_id_from_path(path)}")
        print(f"Это папка объекта: {is_object_folder(path)}")
        files = list_raw_files(path)
        print(f"Сырые файлы ({len(files)}):")
        for f in files:
            print(f"  {f.name}")
