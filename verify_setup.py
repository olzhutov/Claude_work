#!/usr/bin/env python3
"""
Скрипт для проверки установки системы агентов.
Проверяет наличие всех файлов, папок и зависимостей.
"""

import sys
from pathlib import Path
import importlib.util

def check_file(path, name):
    """Проверить наличие файла."""
    if Path(path).exists():
        print(f"✓ {name}")
        return True
    else:
        print(f"✗ {name} (не найден: {path})")
        return False

def check_directory(path, name):
    """Проверить наличие директории."""
    if Path(path).is_dir():
        print(f"✓ {name}")
        return True
    else:
        print(f"✗ {name} (не найдена: {path})")
        return False

def check_module(module_name, package_name=None):
    """Проверить наличие Python модуля."""
    spec = importlib.util.find_spec(module_name)
    if spec is not None:
        print(f"✓ {package_name or module_name}")
        return True
    else:
        print(f"✗ {package_name or module_name} (pip3 install {package_name or module_name})")
        return False

def check_environment():
    """Проверить переменные окружения."""
    import os
    if os.environ.get("ANTHROPIC_API_KEY"):
        print("✓ ANTHROPIC_API_KEY установлена")
        return True
    else:
        print("✗ ANTHROPIC_API_KEY не установлена")
        print("  → export ANTHROPIC_API_KEY=\"sk-ant-...\"")
        return False

def main():
    base_path = Path(__file__).parent
    all_ok = True

    print("\n" + "="*70)
    print("ПРОВЕРКА УСТАНОВКИ СИСТЕМЫ АГЕНТОВ")
    print("="*70 + "\n")

    # Проверка файлов агентов
    print("📁 Файлы агентов:")
    agents = [
        ("agents/__init__.py", "agents/__init__.py"),
        ("agents/config.py", "agents/config.py"),
        ("agents/schemas.py", "agents/schemas.py"),
        ("agents/extractor.py", "agents/extractor.py"),
        ("agents/analyzer.py", "agents/analyzer.py"),
        ("agents/notion_publisher.py", "agents/notion_publisher.py"),
        ("agents/memo_generator.py", "agents/memo_generator.py"),
        ("agents/presentation_builder.py", "agents/presentation_builder.py"),
        ("agents/folder_manager.py", "agents/folder_manager.py"),
        ("agents/pipeline.py", "agents/pipeline.py"),
    ]
    for path, name in agents:
        all_ok &= check_file(base_path / path, name)

    # Проверка основных файлов
    print("\n📄 Основные файлы:")
    files = [
        ("cre_analyzer.py", "cre_analyzer.py"),
        ("CLAUDE.md", "CLAUDE.md"),
        ("AGENTS_README.md", "AGENTS_README.md"),
    ]
    for path, name in files:
        all_ok &= check_file(base_path / path, name)

    # Проверка структуры папок объектов
    print("\n📦 Структура папок объектов:")
    dirs = [
        ("data/objects", "data/objects"),
        ("data/objects/warehouse_kyiv_001", "warehouse_kyiv_001"),
        ("data/objects/warehouse_kyiv_001/raw", "warehouse_kyiv_001/raw"),
        ("data/objects/warehouse_kyiv_001/raw/documents", "warehouse_kyiv_001/raw/documents"),
        ("data/objects/warehouse_kyiv_001/raw/photos", "warehouse_kyiv_001/raw/photos"),
        ("data/objects/warehouse_kyiv_001/raw/plans", "warehouse_kyiv_001/raw/plans"),
        ("data/objects/warehouse_kyiv_001/extracted", "warehouse_kyiv_001/extracted"),
        ("data/objects/warehouse_kyiv_001/output", "warehouse_kyiv_001/output"),
    ]
    for path, name in dirs:
        all_ok &= check_directory(base_path / path, name)

    # Проверка тестового документа
    print("\n📄 Тестовые документы:")
    all_ok &= check_file(
        base_path / "data/objects/warehouse_kyiv_001/raw/documents/test_document.txt",
        "test_document.txt"
    )

    # Проверка зависимостей
    print("\n🔧 Зависимости Python:")
    deps_ok = check_module("anthropic", "anthropic")

    print("\n⚠️  Опциональные зависимости:")
    check_module("PyPDF2", "PyPDF2 (для PDF)")
    check_module("pptx", "python-pptx (для PPTX)")

    # Проверка окружения
    print("\n🔐 Окружение:")
    env_ok = check_environment()

    # Итоги
    print("\n" + "="*70)
    if all_ok and deps_ok and env_ok:
        print("✅ ВСЯ УСТАНОВКА ГОТОВА К ТЕСТИРОВАНИЮ!")
        print("\nЗапустите:")
        print("  python3 agents/pipeline.py data/objects/warehouse_kyiv_001 \\")
        print("    --category income --doc-type memo --format gamma --currency USD")
        return 0
    else:
        print("⚠️  НАЙДЕНЫ ПРОБЛЕМЫ:")
        if not all_ok:
            print("  - Не все файлы на месте")
        if not deps_ok:
            print("  - Не установлена зависимость anthropic")
            print("    → pip3 install anthropic")
        if not env_ok:
            print("  - Не установлена переменная ANTHROPIC_API_KEY")
            print("    → export ANTHROPIC_API_KEY=\"sk-ant-...\"")
        return 1

if __name__ == "__main__":
    sys.exit(main())
