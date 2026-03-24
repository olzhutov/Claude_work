"""
Агент-издатель для Notion: генерирует информационную справку об объекте
и публикует её в Notion через API.
Выходной файл: output/info_brief.md (9 разделов, адаптируется под тип объекта)
"""

import json
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional

from schemas import PropertyData, PipelineConfig


# ---------------------------------------------------------------------------
# Notion API helpers
# ---------------------------------------------------------------------------

NOTION_API_VERSION = "2022-06-28"
NOTION_BASE_URL = "https://api.notion.com/v1"


def _notion_request(
    method: str,
    endpoint: str,
    token: str,
    payload: Optional[Dict] = None,
) -> Dict:
    """Выполняет запрос к Notion API. Возвращает parsed JSON."""
    url = f"{NOTION_BASE_URL}{endpoint}"
    data = json.dumps(payload).encode("utf-8") if payload else None
    req = urllib.request.Request(
        url,
        data=data,
        method=method,
        headers={
            "Authorization": f"Bearer {token}",
            "Notion-Version": NOTION_API_VERSION,
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Notion API {method} {endpoint} → {e.code}: {body}") from e


def _append_blocks(page_id: str, blocks: List[Dict], token: str) -> None:
    """Добавляет блоки на страницу Notion (по 100 за раз — лимит API)."""
    for i in range(0, len(blocks), 100):
        batch = blocks[i : i + 100]
        _notion_request(
            "PATCH",
            f"/blocks/{page_id}/children",
            token,
            {"children": batch},
        )
        if i + 100 < len(blocks):
            time.sleep(0.3)  # избегаем rate limit


# ---------------------------------------------------------------------------
# Markdown → Notion blocks converter
# ---------------------------------------------------------------------------

def _text(content: str, bold: bool = False, code: bool = False) -> Dict:
    """Rich text object для Notion."""
    obj: Dict[str, Any] = {"type": "text", "text": {"content": content}}
    annotations: Dict[str, Any] = {}
    if bold:
        annotations["bold"] = True
    if code:
        annotations["code"] = True
    if annotations:
        obj["annotations"] = annotations
    return obj


def _parse_inline(text: str) -> List[Dict]:
    """Разбирает inline **bold** и `code` в список rich text объектов."""
    import re
    parts = []
    # Разбиваем по **bold** и `code`
    pattern = re.compile(r"(\*\*[^*]+\*\*|`[^`]+`)")
    last = 0
    for m in pattern.finditer(text):
        if m.start() > last:
            parts.append(_text(text[last : m.start()]))
        chunk = m.group(0)
        if chunk.startswith("**"):
            parts.append(_text(chunk[2:-2], bold=True))
        else:
            parts.append(_text(chunk[1:-1], code=True))
        last = m.end()
    if last < len(text):
        parts.append(_text(text[last:]))
    return parts or [_text(text)]


def _md_to_blocks(markdown: str) -> List[Dict]:
    """Конвертирует markdown в список Notion blocks."""
    blocks: List[Dict] = []
    lines = markdown.splitlines()
    i = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Горизонтальный разделитель
        if stripped in ("---", "***", "___"):
            blocks.append({"object": "block", "type": "divider", "divider": {}})
            i += 1
            continue

        # Заголовки
        if stripped.startswith("### "):
            blocks.append({
                "object": "block", "type": "heading_3",
                "heading_3": {"rich_text": _parse_inline(stripped[4:])},
            })
            i += 1
            continue
        if stripped.startswith("## "):
            blocks.append({
                "object": "block", "type": "heading_2",
                "heading_2": {"rich_text": _parse_inline(stripped[3:])},
            })
            i += 1
            continue
        if stripped.startswith("# "):
            blocks.append({
                "object": "block", "type": "heading_1",
                "heading_1": {"rich_text": _parse_inline(stripped[2:])},
            })
            i += 1
            continue

        # Таблица: собираем подряд идущие строки с "|"
        if stripped.startswith("|") and stripped.endswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            # Фильтруем разделитель |---|---|
            rows = [
                r for r in table_lines
                if not all(c in "-| " for c in r)
            ]
            if rows:
                # Разбираем ячейки
                parsed_rows = []
                for row in rows:
                    cells = [c.strip() for c in row.strip("|").split("|")]
                    parsed_rows.append(cells)
                col_count = max(len(r) for r in parsed_rows)
                table_block = {
                    "object": "block",
                    "type": "table",
                    "table": {
                        "table_width": col_count,
                        "has_column_header": True,
                        "has_row_header": False,
                        "children": [],
                    },
                }
                for row_idx, row_cells in enumerate(parsed_rows):
                    cells_blocks = []
                    for cell_text in row_cells:
                        cells_blocks.append(_parse_inline(cell_text))
                    # Дополняем до col_count
                    while len(cells_blocks) < col_count:
                        cells_blocks.append([_text("")])
                    table_block["table"]["children"].append({
                        "type": "table_row",
                        "table_row": {"cells": cells_blocks},
                    })
                blocks.append(table_block)
            continue

        # Маркированный список
        if stripped.startswith("- ") or stripped.startswith("* "):
            content = stripped[2:]
            blocks.append({
                "object": "block", "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": _parse_inline(content)},
            })
            i += 1
            continue

        # Callout (строки с ⚠️)
        if stripped.startswith("> ⚠️") or stripped.startswith("> ⚠"):
            content = stripped.lstrip("> ").strip()
            blocks.append({
                "object": "block", "type": "callout",
                "callout": {
                    "rich_text": _parse_inline(content),
                    "icon": {"type": "emoji", "emoji": "⚠️"},
                    "color": "yellow_background",
                },
            })
            i += 1
            continue

        # Quote (blockquote)
        if stripped.startswith("> "):
            content = stripped[2:]
            blocks.append({
                "object": "block", "type": "quote",
                "quote": {"rich_text": _parse_inline(content)},
            })
            i += 1
            continue

        # Пустая строка → пропускаем (Notion сам добавит отступы)
        if not stripped:
            i += 1
            continue

        # Обычный абзац
        blocks.append({
            "object": "block", "type": "paragraph",
            "paragraph": {"rich_text": _parse_inline(stripped)},
        })
        i += 1

    return blocks


# ---------------------------------------------------------------------------
# Публикация в Notion
# ---------------------------------------------------------------------------

def _get_database(database_id: str, token: str) -> Dict:
    """Загружает схему базы данных Notion (список свойств и их типы)."""
    return _notion_request("GET", f"/databases/{database_id}", token)


def _find_title_prop(schema: Dict) -> str:
    """Находит имя title-свойства в схеме базы данных."""
    for name, prop in schema.items():
        if prop.get("type") == "title":
            return name
    return "Name"


def _build_db_properties(
    property_data: PropertyData,
    metrics: Optional[Dict[str, Any]],
    schema: Dict,
    title: str,
    currency: str,
) -> Dict:
    """
    Формирует словарь properties для записи в базу данных Notion.
    Сопоставляет поля property_data со схемой базы (по имени, без учёта регистра).
    """
    props: Dict[str, Any] = {}

    # Нормализованные имена свойств из схемы → оригинальные имена
    schema_lower = {k.lower(): (k, v) for k, v in schema.items()}

    def _set_prop(search_names: List[str], value: Any) -> None:
        """Ищет подходящее свойство в схеме и устанавливает значение."""
        if value is None:
            return
        for sname in search_names:
            match = schema_lower.get(sname.lower())
            if not match:
                continue
            orig_name, prop_def = match
            ptype = prop_def.get("type")
            if ptype == "title":
                props[orig_name] = {
                    "title": [{"type": "text", "text": {"content": str(value)}}]
                }
            elif ptype == "rich_text":
                props[orig_name] = {
                    "rich_text": [{"type": "text", "text": {"content": str(value)}}]
                }
            elif ptype == "number":
                try:
                    props[orig_name] = {"number": float(value)}
                except (TypeError, ValueError):
                    pass
            elif ptype == "select":
                props[orig_name] = {"select": {"name": str(value)}}
            elif ptype == "url":
                props[orig_name] = {"url": str(value)}
            break

    # 1. Заголовок (title-свойство — обязательно)
    title_prop = _find_title_prop(schema)
    props[title_prop] = {
        "title": [{"type": "text", "text": {"content": title}}]
    }

    # 2. Город
    _set_prop(["city", "місто", "город", "Місто", "City"], property_data.get("city"))

    # 3. Тип объекта
    _set_prop(
        ["type", "тип", "property type", "property_type", "Тип", "Type"],
        property_data.get("property_type"),
    )

    # 4. Адрес
    _set_prop(
        ["address", "адреса", "адрес", "Address", "Адреса"],
        property_data.get("address"),
    )

    # 5. GBA / площадь
    _set_prop(
        ["gba", "площа", "area", "total area", "площадь", "GBA", "Площа"],
        property_data.get("gba"),
    )

    # 6. GLA
    _set_prop(["gla", "GLA", "rentable area"], property_data.get("gla"))

    # 7. Стоимость
    price_val = (
        property_data.get("value")
        or (metrics.get("value") if metrics else None)
    )
    _set_prop(
        ["price", "value", "ціна", "вартість", "Price", "Ціна", "Value"],
        price_val,
    )

    # 8. Финансовые метрики
    if metrics:
        _set_prop(["noi", "NOI"], metrics.get("noi"))
        _set_prop(
            ["cap rate", "cap_rate", "Cap Rate", "Капіталізація"],
            round(metrics.get("cap_rate", 0), 2) if metrics.get("cap_rate") else None,
        )
        _set_prop(
            ["payback", "payback years", "окупність", "Окупність"],
            round(metrics.get("payback_years", 0), 1) if metrics.get("payback_years") else None,
        )

    # 9. Валюта
    _set_prop(
        ["currency", "валюта", "Currency", "Валюта"],
        currency,
    )

    return props


def publish_to_database(
    title: str,
    markdown_content: str,
    token: str,
    database_id: str,
    property_data: Optional[PropertyData] = None,
    metrics: Optional[Dict[str, Any]] = None,
    currency: str = "USD",
) -> str:
    """
    Создаёт запись (страницу) в базе данных Notion.
    Заполняет свойства из property_data (если переданы) и добавляет контент из markdown.
    Возвращает URL созданной записи.
    """
    # Получаем схему базы данных
    try:
        db = _get_database(database_id, token)
        schema = db.get("properties", {})
    except Exception:
        # Если не удалось получить схему — создаём с минимальными свойствами
        schema = {}

    # Формируем свойства записи
    if property_data is not None:
        db_props = _build_db_properties(property_data, metrics, schema, title, currency)
    else:
        # Только заголовок
        title_prop = _find_title_prop(schema)
        db_props = {
            title_prop: {"title": [{"type": "text", "text": {"content": title}}]}
        }

    # Создаём запись в базе данных
    page = _notion_request(
        "POST",
        "/pages",
        token,
        {
            "parent": {"type": "database_id", "database_id": database_id},
            "properties": db_props,
        },
    )
    page_id = page["id"]

    # Добавляем контент как блоки
    blocks = _md_to_blocks(markdown_content)
    _append_blocks(page_id, blocks, token)

    return page.get("url", f"https://notion.so/{page_id.replace('-', '')}")


def publish_to_notion(
    title: str,
    markdown_content: str,
    token: str,
    parent_page_id: str,
) -> str:
    """
    Создаёт страницу в Notion под parent_page_id (page parent), заполняет блоками.
    Возвращает URL созданной страницы.
    """
    page = _notion_request(
        "POST",
        "/pages",
        token,
        {
            "parent": {"type": "page_id", "page_id": parent_page_id},
            "properties": {
                "title": {
                    "title": [{"type": "text", "text": {"content": title}}]
                }
            },
        },
    )
    page_id = page["id"]
    blocks = _md_to_blocks(markdown_content)
    _append_blocks(page_id, blocks, token)
    return page.get("url", f"https://notion.so/{page_id.replace('-', '')}")


# ---------------------------------------------------------------------------
# Основная функция
# ---------------------------------------------------------------------------

def generate_notion_brief(
    property_data: PropertyData,
    metrics: Optional[Dict[str, Any]],
    config: PipelineConfig,
    output_dir: str = "output",
) -> str:
    """
    Генерирует полную информационную справку об объекте (9 разделов),
    сохраняет локально и (если NOTION_TOKEN задан) публикует как запись в БД Notion.

    Args:
        property_data: Извлечённые данные об объекте
        metrics: Финансовые метрики (или None для перспективных объектов)
        config: Конфигурация pipeline
        output_dir: Директория для выходных файлов

    Returns:
        Путь к созданному локальному файлу
    """
    content = _build_info_brief(property_data, metrics, config)

    # Сохраняем локально
    output_path = Path(output_dir) / "info_brief.md"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")

    # Публикуем в Notion если токен задан
    try:
        from config import NOTION_TOKEN, NOTION_DATABASE_ID
        if NOTION_TOKEN and NOTION_DATABASE_ID:
            title = (
                f"{property_data['property_name']} | "
                f"{property_data.get('city', '')} | "
                f"Довідка"
            )
            notion_url = publish_to_database(
                title=title,
                markdown_content=content,
                token=NOTION_TOKEN,
                database_id=NOTION_DATABASE_ID,
                property_data=property_data,
                metrics=metrics,
                currency=config.get("currency", "USD"),
            )
            print(f"  → Notion: {notion_url}")
    except ImportError:
        pass
    except Exception as e:
        print(f"  ⚠ Notion публикація не вдалася: {e}")

    return str(output_path)


# ---------------------------------------------------------------------------
# Построение контента справки (9 разделов)
# ---------------------------------------------------------------------------

def _build_info_brief(
    property_data: PropertyData,
    metrics: Optional[Dict[str, Any]],
    config: PipelineConfig,
) -> str:
    """Собирает полную информационную справку (9 разделов)."""

    lines = []
    currency = config["currency"]
    currency_symbol = "$" if currency == "USD" else "₴"
    prop_type = property_data.get("property_type", "").lower()

    # === РАЗДЕЛ 1: РЕЗЮМЕ ===
    lines.append("# ІНФОРМАЦІЙНА ДОВІДКА ПО ОБ'ЄКТУ")
    lines.append("")
    lines.append("## 1. РЕЗЮМЕ (Executive Summary)")
    lines.append("")
    lines.append(
        f"**Об'єкт:** {property_data['property_name']} | "
        f"**Тип:** {property_data.get('property_type', 'N/A')} | "
        f"**Місто:** {property_data['city']}"
    )
    lines.append("")

    if metrics:
        lines.append("### Ключові показники")
        lines.append(
            f"- **Площа:** {property_data['gba']:,.0f} м² (GBA) / "
            f"{property_data.get('gla', property_data['gba']):,.0f} м² (GLA)"
        )
        lines.append(f"- **Ціна:** {metrics.get('value', 'N/A'):,.0f} {currency_symbol}")
        lines.append(
            f"- **NOI:** {metrics.get('noi', 'N/A'):,.0f} {currency_symbol}/рік"
        )
        lines.append(f"- **Cap Rate:** {metrics.get('cap_rate', 'N/A'):.2f}%")
        lines.append(
            f"- **Окупність:** {metrics.get('payback_years', 'N/A'):.1f} років"
        )
    else:
        lines.append("### Ключові показники")
        lines.append(f"- **Площа:** {property_data['gba']:,.0f} м²")
        lines.append(
            f"- **Ціна:** {property_data.get('value', 'N/A'):,.0f} {currency_symbol}"
        )

    lines.append("")
    lines.append(property_data.get("description", "Опис відсутній"))
    lines.append("")

    # === РАЗДЕЛ 2: ОБЩАЯ ИНФОРМАЦИЯ ===
    lines.append("---")
    lines.append("")
    lines.append("## 2. ЗАГАЛЬНА ІНФОРМАЦІЯ")
    lines.append("")
    lines.append("| Параметр | Значення |")
    lines.append("|----------|----------|")
    lines.append(f"| Назва об'єкту | {property_data['property_name']} |")
    lines.append(f"| Тип / Клас | {property_data.get('property_type', 'N/A')} |")
    if property_data.get("property_class"):
        lines.append(f"| Клас об'єкту | {property_data['property_class']} |")
    lines.append(f"| Адреса | {property_data.get('address', 'N/A')} |")
    lines.append(f"| Місто | {property_data['city']} |")
    if property_data.get("ownership_form"):
        lines.append(f"| Форма власності | {property_data['ownership_form']} |")
    if property_data.get("legal_encumbrances"):
        lines.append(f"| Юридичні обтяження | {property_data['legal_encumbrances']} |")
    if property_data.get("permits_available"):
        permit_status = "Є" if property_data["permits_available"] else "Немає"
        lines.append(f"| Дозвільні документи | {permit_status} |")
    lines.append("")

    # === РАЗДЕЛ 3: ТЕХНИЧЕСКИЕ ХАРАКТЕРИСТИКИ ===
    lines.append("---")
    lines.append("")
    lines.append("## 3. ТЕХНІЧНІ ХАРАКТЕРИСТИКИ")
    lines.append("")
    lines.append("### Площі та розміри")
    lines.append(f"- **GBA (загальна площа):** {property_data['gba']:,.0f} м²")
    if property_data.get("gla"):
        lines.append(f"- **GLA (орендована площа):** {property_data['gla']:,.0f} м²")

    lines.append("")
    lines.append("### Спеціальні характеристики")

    if "склад" in prop_type or "логіст" in prop_type:
        if property_data.get("ceiling_height"):
            lines.append(f"- **Висота стель:** {property_data['ceiling_height']} м")
        if property_data.get("loading_docks"):
            lines.append(f"- **Вантажні дока:** {property_data['loading_docks']} шт")
        if property_data.get("power_capacity_kva"):
            lines.append(f"- **Електромощність:** {property_data['power_capacity_kva']} кВА")

    elif "офіс" in prop_type:
        if property_data.get("ceiling_height"):
            lines.append(f"- **Висота стель:** {property_data['ceiling_height']} м")
        if property_data.get("power_capacity_kva"):
            lines.append(f"- **Електромощність:** {property_data['power_capacity_kva']} кВА")

    elif "виробництво" in prop_type or "виробн" in prop_type:
        if property_data.get("crane_capacity_tons"):
            lines.append(f"- **Кран-балка:** {property_data['crane_capacity_tons']} т")
        if property_data.get("power_capacity_kva"):
            lines.append(f"- **Електромощність:** {property_data['power_capacity_kva']} кВА")

    lines.append(f"- **Рік побудови:** {property_data.get('year_built', 'N/A')}")
    if property_data.get("condition"):
        lines.append(f"- **Стан:** {property_data['condition']}")
    lines.append("")

    # === РАЗДЕЛ 4: ЗЕМЕЛЬНЫЙ УЧАСТОК ===
    lines.append("---")
    lines.append("")
    lines.append("## 4. ЗЕМЕЛЬНА ДІЛЯНКА")
    lines.append("")
    lines.append("| Параметр | Значення |")
    lines.append("|----------|----------|")
    if property_data.get("land_area_ha"):
        lines.append(f"| Площа ділянки | {property_data['land_area_ha']} га |")
    elif property_data.get("land_area"):
        lines.append(f"| Площа ділянки | {property_data['land_area']} га |")
    if property_data.get("land_cadastre"):
        lines.append(f"| Кадастровий номер | {property_data['land_cadastre']} |")
    if property_data.get("land_purpose"):
        lines.append(f"| Цільове призначення | {property_data['land_purpose']} |")
    elif property_data.get("land_category"):
        lines.append(f"| Категорія земель | {property_data['land_category']} |")
    if property_data.get("land_ownership"):
        lines.append(f"| Право на землю | {property_data['land_ownership']} |")
    if property_data.get("land_lease_start") and property_data.get("land_lease_end"):
        lines.append(
            f"| Строк оренди | {property_data['land_lease_start']} – {property_data['land_lease_end']} |"
        )
    elif property_data.get("land_lease_years"):
        lines.append(f"| Строк оренди | {property_data['land_lease_years']} р. |")
    if property_data.get("land_lease_auto_renewal"):
        lines.append("| Автопродовження | Так |")
    lines.append("")

    # === РАЗДЕЛ 5: СОСТОЯНИЕ И ИНФРАСТРУКТУРА ===
    lines.append("---")
    lines.append("")
    lines.append("## 5. СТАН ОБ'ЄКТУ ТА ІНФРАСТРУКТУРА")
    lines.append("")
    if property_data.get("location_description"):
        lines.append("### Локація та доступність")
        lines.append(property_data["location_description"])
        lines.append("")
    if property_data.get("infrastructure"):
        lines.append("### Інженерна інфраструктура")
        infra = property_data["infrastructure"]
        if isinstance(infra, list):
            for item in infra:
                lines.append(f"- {item}")
        else:
            lines.append(str(infra))
        lines.append("")
    if property_data.get("transport_accessibility"):
        lines.append("### Транспортна доступність")
        lines.append(property_data["transport_accessibility"])
        lines.append("")
    if property_data.get("distance_to_highway_km"):
        lines.append(f"**Відстань до КАД:** {property_data['distance_to_highway_km']} км")
        lines.append("")

    # === РАЗДЕЛ 6: РЫНОЧНЫЙ КОНТЕКСТ ===
    lines.append("---")
    lines.append("")
    lines.append("## 6. РИНКОВИЙ КОНТЕКСТ")
    lines.append("")
    if property_data.get("market_rent_rate"):
        lines.append(
            f"**Середня ставка в сегменті:** {property_data['market_rent_rate']:.2f} "
            f"{currency_symbol}/м²/міс"
        )
    if property_data.get("market_vacancy_rate"):
        vacancy_pct = property_data["market_vacancy_rate"] * 100
        lines.append(f"**Вакансія ринку:** {vacancy_pct:.1f}%")
    if property_data.get("market_trends"):
        lines.append("")
        lines.append("### Тренди та прогноз")
        lines.append(property_data["market_trends"])
    if property_data.get("competitors_info"):
        lines.append("")
        lines.append("### Основні конкуренти")
        lines.append(property_data["competitors_info"])
    lines.append("")

    # === РАЗДЕЛ 7: ФИНАНСОВЫЕ ПОКАЗАТЕЛИ ===
    lines.append("---")
    lines.append("")
    if metrics:
        lines.append("## 7. ФІНАНСОВІ ПОКАЗНИКИ")
        lines.append("")
        lines.append("| Метрика | Значення |")
        lines.append("|---------|----------|")
        lines.append(f"| Вартість об'єкту | {metrics['value']:,.0f} {currency_symbol} |")
        if metrics.get("noi"):
            lines.append(
                f"| NOI (чистий дохід) | {metrics['noi']:,.0f} {currency_symbol}/рік |"
            )
        if metrics.get("egi"):
            lines.append(
                f"| EGI (ефективний дохід) | {metrics['egi']:,.0f} {currency_symbol}/рік |"
            )
        if metrics.get("cap_rate"):
            lines.append(f"| Cap Rate | {metrics['cap_rate']:.2f}% |")
        if metrics.get("payback_years"):
            lines.append(f"| Окупність | {metrics['payback_years']:.1f} років |")
        if metrics.get("price_per_unit"):
            unit_name = metrics.get("unit_name", "м²")
            lines.append(
                f"| Ціна за одиницю | {metrics['price_per_unit']:,.0f} "
                f"{currency_symbol}/{unit_name} |"
            )
        if metrics.get("cap_rate_scenarios"):
            lines.append("")
            lines.append("### Сценарії оцінки (за окупністю)")
            lines.append("| Окупність | Вартість | Cap Rate |")
            lines.append("|-----------|----------|----------|")
            for scenario in metrics["cap_rate_scenarios"]:
                if isinstance(scenario, dict):
                    cap_rate = scenario.get("cap_rate_pct")
                    value = scenario.get("property_value")
                    years = scenario.get("payback_years")
                elif isinstance(scenario, (list, tuple)) and len(scenario) >= 2:
                    cap_rate = scenario[0]
                    value = scenario[1]
                    years = None
                else:
                    continue
                if cap_rate and value:
                    years_str = f"{years:.1f} р." if years else "—"
                    lines.append(
                        f"| {years_str} | {value:,.0f} {currency_symbol} | {cap_rate:.1f}% |"
                    )
        lines.append("")
    else:
        lines.append("## 7. ФІНАНСОВІ ПОКАЗНИКИ")
        lines.append("*Фінансовий аналіз не вимагається для перспективних об'єктів*")
        lines.append("")

    # === РАЗДЕЛ 8: ОЦЕНКА РИСКОВ ===
    lines.append("---")
    lines.append("")
    lines.append("## 8. ОЦІНКА РИЗИКІВ")
    lines.append("")
    lines.append("| Категорія | Рівень | Опис |")
    lines.append("|-----------|--------|------|")

    def _risk_row(label: str, risk_text: Optional[str]) -> str:
        if not risk_text:
            return f"| {label} | — | Не оцінено |"
        # Извлекаем уровень из начала строки если есть
        level = "Середній"
        low_markers = ["низький", "низкий", "low"]
        high_markers = ["підвищений", "высокий", "high", "критич"]
        text_lower = risk_text.lower()
        if any(m in text_lower for m in low_markers):
            level = "Низький"
        elif any(m in text_lower for m in high_markers):
            level = "Підвищений"
        return f"| {label} | {level} | {risk_text} |"

    lines.append(_risk_row("Юридичні", property_data.get("risk_legal")))
    lines.append(_risk_row("Ринкові", property_data.get("risk_market")))
    lines.append(_risk_row("Технічні", property_data.get("risk_technical")))
    lines.append("")

    # === РАЗДЕЛ 9: ДОКУМЕНТЫ И ПРИЛОЖЕНИЯ ===
    lines.append("---")
    lines.append("")
    lines.append("## 9. ДОКУМЕНТИ ТА ДОДАТКИ")
    lines.append("")

    if property_data.get("documents_list"):
        lines.append("### Наявні документи")
        for doc in property_data["documents_list"]:
            lines.append(f"- {doc}")
        lines.append("")

    if property_data.get("photos_count"):
        lines.append(f"**Фотографії:** {property_data['photos_count']} шт.")
    if property_data.get("plans_available"):
        lines.append("**Планування/креслення:** Є")

    if property_data.get("extraction_notes"):
        lines.append("")
        lines.append("### Примітки щодо даних")
        lines.append(property_data["extraction_notes"])

    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append(
        f"*Довідка сформована автоматично на підставі: "
        f"{property_data.get('source_file', 'документів об\u2019єкта')}*"
    )

    return "\n".join(lines)


if __name__ == "__main__":
    pass
