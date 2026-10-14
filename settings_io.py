"""
Экспорт и импорт настроек приложения (settings.db) в/из Excel-файла.

Формат файла — одна книга со следующими листами:
  1. "Поставщики"         — supplier_brands (ИНН → магазин)
  2. "Включения продуктов" — food_item_includes
  3. "Подкатегории продуктов" — food_categories (Огурцы → Овощи)
  4. "Категории"          — expense_categories (Название, Порядок)
  5. "Магазины категорий" — expense_category_stores (Категория, Магазин)
  6. "Исключения категорий" — expense_category_item_excludes
                               (Категория, Магазин, Подстрока товара)

При импорте — полная замена содержимого каждой таблицы. Это безопаснее,
чем merge, и позволяет удалять ненужные записи прямо в Excel.
"""
from __future__ import annotations

import logging
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

logger = logging.getLogger(__name__)


HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2B5797", end_color="2B5797",
                          fill_type="solid")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center",
                             wrap_text=True)
THIN = Side(style="thin", color="B0B0B0")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


SHEET_SUPPLIERS = "Поставщики"
SHEET_FOOD_INCLUDES = "Включения продуктов"
SHEET_FOOD_SUBCATS = "Подкатегории продуктов"
SHEET_CATEGORIES = "Категории"
SHEET_CATEGORY_STORES = "Магазины категорий"
SHEET_CATEGORY_EXCLUDES = "Исключения категорий"


def _apply_header(ws, row: int, col_count: int):
    for c in range(1, col_count + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER


def _autosize(ws, widths: list[int]):
    from openpyxl.utils import get_column_letter
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_sheet(ws, headers: list[str], rows: list[list],
                 widths: list[int]):
    for ci, h in enumerate(headers):
        ws.cell(row=1, column=ci + 1, value=h)
    _apply_header(ws, 1, len(headers))

    for ri, row in enumerate(rows, 2):
        for ci, val in enumerate(row):
            ws.cell(row=ri, column=ci + 1, value=val).border = THIN_BORDER

    _autosize(ws, widths)


def export_settings_to_excel(settings_db, output_path: str) -> str:
    """Выгрузить все настройки в один Excel-файл."""
    wb = Workbook()
    wb.remove(wb.active)  # удалить дефолтный лист

    # 1) Поставщики (supplier_brands)
    ws = wb.create_sheet(SHEET_SUPPLIERS)
    brands = settings_db.get_all_supplier_brands()
    rows = [[b.get("provider_inn", ""), b.get("brand_name", ""),
             b.get("description", "") or ""] for b in brands]
    _write_sheet(ws, ["ИНН поставщика", "Магазин", "Описание"],
                 rows, [18, 35, 40])

    # 2) Включения продуктов (food_item_includes)
    ws = wb.create_sheet(SHEET_FOOD_INCLUDES)
    includes = settings_db.get_food_item_includes()
    rows = [[i.get("item_pattern", ""), i.get("description", "") or ""]
            for i in includes]
    _write_sheet(ws, ["Подстрока товара", "Описание"], rows, [40, 40])

    # 3) Подкатегории продуктов (food_categories: огурцы → овощи)
    ws = wb.create_sheet(SHEET_FOOD_SUBCATS)
    subcats = settings_db.get_all_food_categories()
    rows = [[s.get("item_pattern", ""), s.get("category", "")]
            for s in subcats]
    _write_sheet(ws, ["Подстрока товара", "Подкатегория"], rows, [40, 30])

    # 4) Категории расходов (expense_categories)
    ws = wb.create_sheet(SHEET_CATEGORIES)
    cats = settings_db.get_expense_categories()
    rows = [[c.get("name", ""), c.get("display_order", 0)] for c in cats]
    _write_sheet(ws, ["Название категории", "Порядок"], rows, [28, 12])

    # 5) Магазины категорий
    ws = wb.create_sheet(SHEET_CATEGORY_STORES)
    cat_stores = settings_db.get_all_category_stores()
    rows = []
    for cat_name, stores in cat_stores.items():
        for s in stores:
            rows.append([cat_name, s])
    _write_sheet(ws, ["Категория", "Магазин"], rows, [28, 35])

    # 6) Исключения товаров категорий
    ws = wb.create_sheet(SHEET_CATEGORY_EXCLUDES)
    excludes = settings_db.get_all_category_item_excludes()
    rows = [[e.get("category", ""), e.get("store_name", "") or "",
             e.get("item_pattern", "")] for e in excludes]
    _write_sheet(ws,
                 ["Категория", "Магазин (опционально)", "Подстрока товара"],
                 rows, [28, 30, 40])

    wb.save(output_path)
    return output_path


def _read_sheet_rows(wb, sheet_name: str) -> list[list]:
    """Прочитать все строки листа (пропуская строку заголовка)."""
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Пропускаем полностью пустые строки
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        rows.append([("" if v is None else str(v).strip()) for v in row])
    return rows


def import_settings_from_excel(settings_db, input_path: str) -> dict:
    """Загрузить настройки из Excel-файла.

    Стратегия: для каждой таблицы — полная замена (сначала очистка,
    потом запись). Это позволяет удалять записи прямо в Excel.

    Возвращает {table_name: count} — количество импортированных записей.
    """
    wb = load_workbook(input_path, read_only=True, data_only=True)
    stats = {}

    # 1) Поставщики
    settings_db.clear_supplier_brands()
    rows = _read_sheet_rows(wb, SHEET_SUPPLIERS)
    count = 0
    for r in rows:
        inn = r[0] if len(r) > 0 else ""
        brand = r[1] if len(r) > 1 else ""
        desc = r[2] if len(r) > 2 else ""
        if inn and brand:
            settings_db.set_supplier_brand(inn, brand, desc)
            count += 1
    stats[SHEET_SUPPLIERS] = count

    # 2) Включения продуктов
    settings_db.clear_food_item_includes()
    rows = _read_sheet_rows(wb, SHEET_FOOD_INCLUDES)
    count = 0
    for r in rows:
        pat = r[0] if len(r) > 0 else ""
        desc = r[1] if len(r) > 1 else ""
        if pat:
            settings_db.add_food_item_include(pat, desc)
            count += 1
    stats[SHEET_FOOD_INCLUDES] = count

    # 3) Подкатегории продуктов
    settings_db.clear_food_categories()
    rows = _read_sheet_rows(wb, SHEET_FOOD_SUBCATS)
    count = 0
    for r in rows:
        pat = r[0] if len(r) > 0 else ""
        cat = r[1] if len(r) > 1 else ""
        if pat and cat:
            settings_db.set_food_category(pat, cat)
            count += 1
    stats[SHEET_FOOD_SUBCATS] = count

    # 4-6) Категории расходов — полная замена блоком, чтобы сохранить
    #      целостность foreign key.
    settings_db.clear_all_expense_categories()

    # 4) Категории (создаём все категории с порядком)
    rows = _read_sheet_rows(wb, SHEET_CATEGORIES)
    count = 0
    for r in rows:
        name = r[0] if len(r) > 0 else ""
        try:
            order = int(r[1]) if len(r) > 1 and r[1] else 0
        except ValueError:
            order = 0
        if name:
            settings_db.add_expense_category(name, order)
            count += 1
    stats[SHEET_CATEGORIES] = count

    # 5) Магазины категорий
    rows = _read_sheet_rows(wb, SHEET_CATEGORY_STORES)
    count = 0
    for r in rows:
        cat = r[0] if len(r) > 0 else ""
        store = r[1] if len(r) > 1 else ""
        if cat and store:
            settings_db.add_category_store(cat, store)
            count += 1
    stats[SHEET_CATEGORY_STORES] = count

    # 6) Исключения товаров категорий
    rows = _read_sheet_rows(wb, SHEET_CATEGORY_EXCLUDES)
    count = 0
    for r in rows:
        cat = r[0] if len(r) > 0 else ""
        store = r[1] if len(r) > 1 else ""
        pat = r[2] if len(r) > 2 else ""
        if cat and pat:
            settings_db.add_category_item_exclude(cat, store, pat)
            count += 1
    stats[SHEET_CATEGORY_EXCLUDES] = count

    wb.close()
    return stats
