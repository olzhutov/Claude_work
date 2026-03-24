#!/usr/bin/env python3
"""
Оркестратор pipeline: главная точка входа для обработки документов.
Запуск: python3 agents/pipeline.py /path/to/document.pdf --category income --doc-type full --format pptx --exchange-rate 41.5 --currency USD
"""

import argparse
import json
import sys
from pathlib import Path

from config import OUTPUT_DIR
from extractor import extract_property_data, extract_from_folder
from analyzer import analyze_property
from notion_publisher import generate_notion_brief
from memo_generator import generate_memo
from presentation_builder import build_presentation
from folder_manager import (
    is_object_folder,
    get_object_id_from_path,
    save_extracted_data,
    load_extracted_data,
)
from schemas import PipelineConfig


def main():
    """Главная функция pipeline."""
    parser = argparse.ArgumentParser(
        description="Обработка документов коммерческой недвижимости"
    )

    parser.add_argument(
        "document",
        help="Путь к исходному документу (PDF, изображение, текст) или папке объекта (структура data/objects/{id}/)",
    )
    parser.add_argument(
        "--category",
        default="income",
        choices=["income", "prospect"],
        help="Категория объекта: income (доходный) или prospect (перспективный)",
    )
    parser.add_argument(
        "--doc-type",
        default="memo",
        choices=["teaser", "memo", "full"],
        help="Тип документа: teaser (3-5 стр), memo (5-10 стр), full (10+ стр)",
    )
    parser.add_argument(
        "--format",
        default="pptx",
        choices=["pptx", "gamma"],
        help="Формат презентации: pptx (PowerPoint) или gamma (текстовый аутлайн)",
    )
    parser.add_argument(
        "--exchange-rate",
        type=float,
        default=41.5,
        help="Курс UAH/USD (дефолт: 41.5)",
    )
    parser.add_argument(
        "--currency",
        default="USD",
        choices=["USD", "UAH"],
        help="Базовая валюта (дефолт: USD)",
    )
    parser.add_argument(
        "--output-dir",
        default=OUTPUT_DIR,
        help=f"Директория для выходных файлов (дефолт: {OUTPUT_DIR})",
    )
    parser.add_argument(
        "--location-score",
        action="store_true",
        default=False,
        help="Выполнить оценку локации и добавить в отчёт",
    )

    args = parser.parse_args()

    # Проверяем что файл / папка существует
    document_path = Path(args.document)
    if not document_path.exists():
        print(f"Ошибка: файл / папка не найдены: {args.document}", file=sys.stderr)
        sys.exit(1)

    # Определяем тип входа: папка объекта или файл
    is_folder = document_path.is_dir() and is_object_folder(str(document_path))

    print("=" * 70)
    print("ОБРАБОТКА КОММЕРЧЕСКОЙ НЕДВИЖИМОСТИ")
    print("=" * 70)

    if is_folder:
        object_id = get_object_id_from_path(str(document_path))
        print(f"Папка об'єкту: {object_id}")
        output_dir_path = document_path / "output"
    else:
        print(f"Документ: {document_path.name}")
        output_dir_path = Path(args.output_dir)

    print(f"Категорія: {args.category}")
    print(f"Тип: {args.doc_type}")
    print(f"Формат: {args.format}")
    print(f"Валюта: {args.currency}")
    if args.location_score:
        print(f"Оцінка локації: увімкнена")
    print()

    try:
        # Шаг 1: Извлечение данных
        if is_folder:
            print("[1/5] Извлечение данных из папки об'єкту...")
            property_data = extract_from_folder(str(document_path))
            # Сохраняем извлечённые данные
            save_extracted_data(str(document_path), property_data)
        else:
            print("[1/5] Извлечение данных из документа...")
            property_data = extract_property_data(str(document_path))

        print(f"✓ Об'єкт: {property_data['property_name']}")
        print(f"✓ Площадь: {property_data['gba']:,.0f} м²")
        print(f"✓ Уверенность: {property_data['extraction_confidence']}")
        print()

        # Шаг 2: Финансовый анализ (если доходный объект)
        config = PipelineConfig(
            object_category=args.category,
            doc_type=args.doc_type,
            pres_format=args.format,
            exchange_rate=args.exchange_rate,
            currency=args.currency,
        )

        metrics = None
        if args.category == "income":
            print("[2/5] Финансовый анализ объекта...")
            metrics = analyze_property(property_data, config)
            if metrics:
                print(f"✓ NOI: {metrics.get('noi', 'N/A'):,.0f} {args.currency}/год")
                print(f"✓ Cap Rate: {metrics.get('cap_rate', 'N/A'):.2f}%")
                print(f"✓ Окупаемость: {metrics.get('payback_years', 'N/A'):.1f} лет")
            print()
        else:
            print("[2/5] Финансовый анализ пропущен (перспективный объект)")
            print()

        # Шаг 3: Генерация информационной справки
        print("[3/5] Генерация інформаційної довідки (Notion)...")
        notion_file = generate_notion_brief(
            property_data, metrics, config, str(output_dir_path)
        )
        print(f"✓ Создан: {notion_file}")
        print()

        # Шаг 4: Генерация меморандума
        print("[4/5] Генерация інвестиційного меморандуму...")
        memo_file = generate_memo(property_data, metrics, config, str(output_dir_path))
        print(f"✓ Создан: {memo_file}")
        print()

        # Шаг 5: Генерация презентации
        print("[5/5] Генерация презентації...")
        pres_file = build_presentation(
            property_data, metrics, config, str(output_dir_path)
        )
        print(f"✓ Создан: {pres_file}")
        print()

        # Шаг 6 (опционально): Оценка локации
        location_score = None
        if args.location_score:
            print("[6/6] Оцінка локації...")
            try:
                from location_scorer import score_location
                address = property_data.get("address", "") + ", " + property_data.get("city", "")
                property_type = property_data.get("property_type", None)
                location_score = score_location(address.strip(", "), property_type)
                # Сохраняем отчёт об оценке локации
                location_file = output_dir_path / "location_score.txt"
                output_dir_path.mkdir(parents=True, exist_ok=True)
                location_file.write_text(location_score["report"], encoding="utf-8")
                print(f"✓ Оцінка локації: {location_score['total_score']}/10 {location_score['status']}")
                print(f"✓ Збережено: {location_file.name}")
            except Exception as e:
                print(f"⚠ Оцінку локації не вдалося виконати: {e}")
            print()

        # Итоги
        print("=" * 70)
        print("ЗАВЕРШЕНО УСПІШНО!")
        print("=" * 70)
        print()
        print("Створені файли:")
        print(f"1. Інформ. довідка:  {Path(notion_file).name}")
        print(f"2. Меморандум:       {Path(memo_file).name}")
        print(f"3. Презентація:      {Path(pres_file).name}")
        if location_score:
            print(f"4. Оцінка локації:   location_score.txt")
        print()

        if args.format == "gamma":
            print("УВАГА: Для Gamma презентації:")
            print(f"  1. Відкрийте файл:  {pres_file}")
            print("  2. Скопіюйте вміст (Ctrl+A, Ctrl+C)")
            print("  3. Вставте в Claude Code")
            print("  4. Попросите створити Gamma-презентацію")
            print()

        if args.category == "income" and metrics:
            print("Фінансові показники:")
            print(f"  • NOI:              {metrics.get('noi', 'N/A'):,.0f} {args.currency}/рік")
            print(f"  • Cap Rate:         {metrics.get('cap_rate', 'N/A'):.2f}%")
            print(f"  • Окупаємість:      {metrics.get('payback_years', 'N/A'):.1f} років")
            print()

        if location_score:
            print("Оцінка локації:")
            print(f"  • Загальна оцінка:  {location_score['total_score']}/10  {location_score['status']}")
            for factor, score in location_score["factor_scores"].items():
                print(f"  • {factor:20s}: {score:.1f}/10")
            print()

        print("Шлях до файлів:")
        print(f"  {output_dir_path.absolute()}")
        print()

    except FileNotFoundError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"Ошибка обработки: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Неожиданная ошибка: {e}", file=sys.stderr)
        import traceback

        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
