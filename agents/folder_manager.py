"""
Менеджер структуры папок объекта.
Управляет созданием и навигацией по структуре: raw/ → extracted/ → output/
"""

import json
from pathlib import Path
from typing import List, Dict, Any, Optional


def create_object_folder(object_id: str, base_path: str = "data/objects") -> Path:
    """
    Создаёт структуру папок для объекта недвижимости.

    Args:
        object_id: ID объекта (например, "warehouse_kyiv_001")
        base_path: Базовая директория для объектов

    Returns:
        Path к созданной папке объекта

    Структура:
        base_path/
        └── object_id/
            ├── raw/
            │   ├── documents/
            │   ├── photos/
            │   └── plans/
            ├── extracted/
            └── output/
    """

    object_dir = Path(base_path) / object_id

    # Создаём все поддиректории
    (object_dir / "raw" / "documents").mkdir(parents=True, exist_ok=True)
    (object_dir / "raw" / "photos").mkdir(parents=True, exist_ok=True)
    (object_dir / "raw" / "plans").mkdir(parents=True, exist_ok=True)
    (object_dir / "extracted").mkdir(parents=True, exist_ok=True)
    (object_dir / "output").mkdir(parents=True, exist_ok=True)

    return object_dir


def is_object_folder(folder_path: str) -> bool:
    """
    Проверяет что папка имеет структуру объекта (есть raw/, extracted/, output/).
    """
    folder = Path(folder_path)
    return (
        (folder / "raw").is_dir()
        and (folder / "extracted").is_dir()
        and (folder / "output").is_dir()
    )


def list_raw_files(
    folder_path: str, subfolder: Optional[str] = None
) -> List[Path]:
    """
    Возвращает список файлов из raw/ директории.

    Args:
        folder_path: Путь к папке объекта
        subfolder: Опциональная поддиректория (documents, photos, plans)

    Returns:
        Список Path объектов к файлам
    """
    folder = Path(folder_path)

    if subfolder:
        raw_dir = folder / "raw" / subfolder
    else:
        raw_dir = folder / "raw"

    if not raw_dir.exists():
        return []

    # Возвращаем все файлы рекурсивно
    files = []
    for file_path in raw_dir.rglob("*"):
        if file_path.is_file():
            files.append(file_path)

    return sorted(files)


def get_extracted_path(folder_path: str) -> Path:
    """Возвращает путь к файлу property_data.json."""
    folder = Path(folder_path)
    return folder / "extracted" / "property_data.json"


def save_extracted_data(folder_path: str, data: Dict[str, Any]) -> Path:
    """
    Сохраняет извлечённые данные в JSON файл.

    Args:
        folder_path: Путь к папке объекта
        data: Словарь с данными PropertyData

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
    Загружает извлечённые данные из JSON файла.

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
    Возвращает путь для файла в output/.

    Args:
        folder_path: Путь к папке объекта
        filename: Имя выходного файла (например, "info_brief.md")

    Returns:
        Path к файлу в output/
    """
    folder = Path(folder_path)
    return folder / "output" / filename


def save_output_file(folder_path: str, filename: str, content: str) -> Path:
    """
    Сохраняет выходной файл.

    Args:
        folder_path: Путь к папке объекта
        filename: Имя файла
        content: Содержимое

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
    Извлекает ID объекта из пути.

    Примеры:
        "data/objects/warehouse_kyiv_001" → "warehouse_kyiv_001"
        "/full/path/to/warehouse_kyiv_001" → "warehouse_kyiv_001"
    """
    folder = Path(folder_path)
    return folder.name


if __name__ == "__main__":
    # Пример использования
    obj_path = create_object_folder("warehouse_test_001")
    print(f"Создана папка: {obj_path}")

    # Список файлов в documents/
    docs = list_raw_files(str(obj_path), "documents")
    print(f"Файлы в documents/: {docs}")

    # Сохранить данные
    test_data = {"property_name": "Тест", "gba": 5000}
    saved_path = save_extracted_data(str(obj_path), test_data)
    print(f"Сохранены данные в: {saved_path}")

    # Загрузить данные
    loaded_data = load_extracted_data(str(obj_path))
    print(f"Загруженные данные: {loaded_data}")
