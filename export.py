"""
Модуль экспорта чеков в Excel с разбивкой по магазинам/поставщикам.
"""
import os
import re
import sys
import logging
from datetime import datetime
from typing import Optional

from export_period import filter_by_years

from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.plotarea import DataTable
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, quote_sheetname


def _resource_path(rel_path: str) -> str:
    """Возвращает путь к ресурсу, работает и для dev, и для PyInstaller exe."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), rel_path)


# Путь к шаблону с настроенным листом "Продукты" (диаграммы, стили)
FOOD_TEMPLATE_PATH = _resource_path("food_template.xlsx")

logger = logging.getLogger(__name__)

# Стили
HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2B5797", end_color="2B5797", fill_type="solid")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)

LINK_FONT = Font(name="Calibri", size=11, color="2B5797", underline="single")

SELLER_FONT = Font(name="Calibri", bold=True, size=12, color="2B5797")
SELLER_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")

RECEIPT_HEADER_FONT = Font(name="Calibri", bold=True, size=10, color="2B5797")
RECEIPT_HEADER_FILL = PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid")

TOTAL_FONT = Font(name="Calibri", bold=True, size=11)
TOTAL_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

MONEY_FORMAT = '#,##0.00 "₽"'

RECEIPT_HEADERS = [
    "Дата", "Номер ФД", "Торговая точка", "Сумма", "Адрес", "ИНН продавца"
]

ITEM_HEADERS = [
    "Дата чека", "Продавец", "Торговая точка", "Наименование товара",
    "Цена", "Кол-во", "Сумма"
]


def _normalize_seller_name(name: str) -> str:
    """Нормализовать имя продавца для группировки (регистронезависимо)."""
    if not name:
        return "Неизвестный продавец"
    # Приводим к верхнему регистру для сравнения
    return name.strip()


def _seller_group_key(name: str) -> str:
    """Ключ группировки продавца (верхний регистр без лишних пробелов)."""
    if not name:
        return ""
    return re.sub(r'\s+', ' ', name.strip().upper())


def _store_name(receipt: dict) -> str:
    """Получить название магазина: brand_name → retail_place → пусто."""
    return (receipt.get("brand_name") or receipt.get("retail_place") or "")


def _abbreviate_org_form(name: str) -> str:
    """Сокращение организационно-правовых форм в названии."""
    replacements = [
        (r'(?i)ОБЩЕСТВО\s+С\s+ОГРАНИЧЕННОЙ\s+ОТВЕТСТВЕННОСТЬЮ', 'ООО'),
        (r'(?i)АКЦИОНЕРНОЕ\s+ОБЩЕСТВО', 'АО'),
        (r'(?i)ПУБЛИЧНОЕ\s+АКЦИОНЕРНОЕ\s+ОБЩЕСТВО', 'ПАО'),
        (r'(?i)ЗАКРЫТОЕ\s+АКЦИОНЕРНОЕ\s+ОБЩЕСТВО', 'ЗАО'),
        (r'(?i)ИНДИВИДУАЛЬНЫЙ\s+ПРЕДПРИНИМАТЕЛЬ', 'ИП'),
    ]
    for pattern, repl in replacements:
        name = re.sub(pattern, repl, name)
    return name


def _safe_sheet_name(name: str, max_len: int = 31) -> str:
    """Создать безопасное имя листа для Excel."""
    # Сокращаем орг. формы
    name = _abbreviate_org_form(name)
    # Убираем ВСЕ символы, запрещённые в именах листов Excel
    for ch in ['\\', '/', '*', '?', ':', '[', ']', '"', "'", '\x00']:
        name = name.replace(ch, '')
    name = re.sub(r'\s+', ' ', name.strip())
    if len(name) > max_len:
        name = name[:max_len - 2] + ".."
    return name or "Лист"


def _unique_sheet_name(wb, desired_name: str) -> str:
    """Получить уникальное имя листа, добавляя суффикс при конфликте."""
    base = _safe_sheet_name(desired_name)
    existing = {ws.title for ws in wb.worksheets}
    if base not in existing:
        return base
    # Пробуем суффиксы _2, _3 и т.д.
    for i in range(2, 100):
        suffix = f"_{i}"
        candidate = _safe_sheet_name(desired_name, 31 - len(suffix)) + suffix
        if candidate not in existing:
            return candidate
    return base[:25] + f"_{id(desired_name) % 10000}"


def _apply_header_style(ws, row: int, col_count: int):
    """Применить стиль к строке заголовков."""
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER


def _auto_width(ws, min_width: int = 10, max_width: int = 50):
    """Автоподбор ширины столбцов."""
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            try:
                val = str(cell.value) if cell.value else ""
                max_len = max(max_len, len(val))
            except Exception:
                pass
        adjusted = min(max(max_len + 2, min_width), max_width)
        ws.column_dimensions[col_letter].width = adjusted


def _format_date(date_str) -> str:
    """Форматировать дату из различных форматов."""
    try:
        if "T" in str(date_str):
            dt = datetime.fromisoformat(str(date_str).replace("Z", "+00:00"))
            return dt.strftime("%d.%m.%Y")
        else:
            return str(date_str)[:10]
    except Exception:
        return str(date_str)


def export_receipts_to_excel(
    receipts: list[dict],
    db,
    output_path: str,
    group_by_seller: bool = True,
    include_items: bool = True,
    years=None,
) -> str:
    """
    Экспортировать чеки в Excel файл.
    """
    receipts = filter_by_years(receipts, years)
    if years is not None and not years:
        raise ValueError("Выберите хотя бы один год для экспорта.")

    # Стартуем с шаблона, где уже настроен лист "Продукты" с диаграммами.
    # Если шаблон недоступен, падаем обратно на пустую книгу.
    if os.path.exists(FOOD_TEMPLATE_PATH):
        wb = load_workbook(FOOD_TEMPLATE_PATH)
        # Добавляем лист "Сводка" в начало
        ws_summary = wb.create_sheet("Сводка", 0)
    else:
        wb = Workbook()
        ws_summary = wb.active
        ws_summary.title = "Сводка"

    ws_summary.merge_cells("A1:F1")
    title_cell = ws_summary["A1"]
    title_cell.value = "Выгрузка чеков — Мои чеки онлайн"
    title_cell.font = Font(name="Calibri", bold=True, size=14, color="2B5797")
    title_cell.alignment = Alignment(horizontal="center")

    ws_summary["A2"] = f"Дата выгрузки: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    ws_summary["A3"] = f"Всего чеков: {len(receipts)}"

    total_amount = sum(r.get("total_sum", 0) or 0 for r in receipts)
    ws_summary["A4"] = f"Общая сумма: {total_amount:,.2f} ₽"
    ws_summary["A4"].font = Font(bold=True, size=12)

    ws_summary["A5"] = "Годы: " + (", ".join(map(str, sorted(set(years)))) if years is not None else "вся история")

    # Группировка по продавцам (регистронезависимо по ИНН, потом по имени)
    seller_by_inn = {}
    for r in receipts:
        inn = (r.get("seller_inn") or "").strip()
        seller_name = _normalize_seller_name(r.get("seller_name", ""))
        group_key = inn if inn else _seller_group_key(seller_name)

        if group_key not in seller_by_inn:
            seller_by_inn[group_key] = {
                "display_name": seller_name,
                "inn": inn,
                "count": 0,
                "total": 0,
                "places": set(),
                "brands": set(),
                "addresses": set(),
            }
        info = seller_by_inn[group_key]
        info["count"] += 1
        info["total"] += r.get("total_sum", 0) or 0
        rp = r.get("retail_place", "")
        if rp:
            info["places"].add(rp)
        bn = r.get("brand_name", "")
        if bn:
            info["brands"].add(bn)
        addr = r.get("seller_address", "")
        if addr:
            info["addresses"].add(addr)
        if len(seller_name) > len(info["display_name"]):
            info["display_name"] = seller_name

    # Сводная таблица по продавцам
    row = 6
    summary_headers = ["Продавец (юр. лицо)", "ИНН", "Магазины",
                        "Чеков", "Сумма", "→ Лист"]
    for col, h in enumerate(summary_headers, 1):
        ws_summary.cell(row=row, column=col, value=h)
    _apply_header_style(ws_summary, row, len(summary_headers))

    # Сортируем по сумме (наибольшая сначала)
    sorted_groups = sorted(
        seller_by_inn.values(),
        key=lambda x: x["total"],
        reverse=True,
    )

    # Словарь для связи группы с именем листа (заполнится позже)
    group_to_sheet: dict[str, str] = {}  # group_key -> sheet_name
    summary_link_rows: list[tuple[int, str]] = []  # (row, group_key)

    for info in sorted_groups:
        row += 1
        ws_summary.cell(row=row, column=1, value=_abbreviate_org_form(info["display_name"]))
        ws_summary.cell(row=row, column=2, value=info["inn"])
        # Показываем бренды (магазины) если есть, иначе retailPlace
        if info["brands"]:
            places_str = ", ".join(sorted(info["brands"]))
        elif info["places"]:
            places_str = ", ".join(sorted(info["places"]))
        else:
            places_str = "—"
        ws_summary.cell(row=row, column=3, value=places_str)
        ws_summary.cell(row=row, column=4, value=info["count"])
        cell_sum = ws_summary.cell(row=row, column=5, value=info["total"])
        cell_sum.number_format = MONEY_FORMAT
        # Ссылку заполним после создания листов
        gk = info["inn"] if info["inn"] else _seller_group_key(info["display_name"])
        summary_link_rows.append((row, gk))
        for col in range(1, len(summary_headers) + 1):
            ws_summary.cell(row=row, column=col).border = THIN_BORDER

    # Итого
    row += 1
    ws_summary.cell(row=row, column=1, value="ИТОГО").font = TOTAL_FONT
    ws_summary.cell(row=row, column=4, value=len(receipts)).font = TOTAL_FONT
    total_cell = ws_summary.cell(row=row, column=5, value=total_amount)
    total_cell.font = TOTAL_FONT
    total_cell.number_format = MONEY_FORMAT
    for col in range(1, len(summary_headers) + 1):
        ws_summary.cell(row=row, column=col).fill = TOTAL_FILL
        ws_summary.cell(row=row, column=col).border = THIN_BORDER

    # ══════════════════════════════════════════════════════════
    # Порядок листов:
    # 1) Сводка (уже создана)
    # 2) Статистика (матрица категория × месяц/год)
    # 3) Продукты (спец-лист с дашбордом)
    # 4) Продукты_данные
    # 5) Другие категории (лист на каждую)
    # 6) Все чеки
    # 7) Все товары
    # 8) Листы по продавцам
    # ══════════════════════════════════════════════════════════

    # --- Данные по категориям ---
    expense_data = None
    if db and hasattr(db, "get_expense_data"):
        try:
            expense_data = db.get_expense_data()
            expense_data = dict(expense_data)
            expense_data["by_category"] = {
                name: filter_by_years(items, years)
                for name, items in expense_data.get("by_category", {}).items()
            }
        except Exception as e:
            logger.warning("Не удалось получить данные категорий: %s", e)

    # --- Лист "Статистика" (2-й лист) ---
    # Создаём сразу после "Сводка", до шаблонного листа "Продукты"
    if expense_data and expense_data.get("by_category"):
        # Индекс листа "Продукты" (если он есть), иначе в конец
        prod_idx = (wb.sheetnames.index("Продукты")
                    if "Продукты" in wb.sheetnames else len(wb.sheetnames))
        ws_stats = wb.create_sheet("Статистика", prod_idx)
        _write_statistics_sheet(ws_stats, expense_data["by_category"])

    # --- Дашборд "Продукты" (спец-лист из шаблона) ---
    if db:
        food_data = filter_by_years(db.get_food_dashboard_data(), years)
        if food_data and "Продукты" not in wb.sheetnames:
            wb.create_sheet("Продукты")
        if food_data:
            _write_food_dashboard(wb, food_data)
            # Лист с исходными данными для проверки (сразу после "Продукты")
            prod_idx = wb.sheetnames.index("Продукты") if "Продукты" in wb.sheetnames else 1
            ws_fd = wb.create_sheet("Продукты_данные", prod_idx + 1)
            _write_food_data_sheet(ws_fd, food_data)

    # Удаляем пустой шаблон: в нём могут оставаться данные старых лет.
    if "Продукты" in wb.sheetnames and (not db or not food_data):
        wb.remove(wb["Продукты"])

    # --- Листы остальных категорий (после Продукты_данные) ---
    if expense_data and expense_data.get("by_category"):
        # Порядок категорий из settings_db (display_order)
        cat_order: list[str] = []
        try:
            if db and db.settings:
                cat_order = db.settings.get_expense_category_names()
        except Exception:
            pass

        insert_after = ("Продукты_данные" if "Продукты_данные" in wb.sheetnames
                        else ("Продукты" if "Продукты" in wb.sheetnames else "Сводка"))
        insert_idx = wb.sheetnames.index(insert_after) + 1

        for cat_name in cat_order:
            if cat_name == "Продукты":
                continue  # у Продуктов отдельный дашборд
            items = expense_data["by_category"].get(cat_name, [])
            if not items:
                continue
            sheet_name = _unique_sheet_name(wb, cat_name)
            ws_cat = wb.create_sheet(sheet_name, insert_idx)
            _write_category_sheet(ws_cat, cat_name, items)
            insert_idx += 1

    # --- Лист "Все чеки" ---
    ws_all = wb.create_sheet("Все чеки")
    _write_receipts_sheet(ws_all, receipts)

    # --- Лист "Все товары" ---
    if include_items and db:
        ws_items = wb.create_sheet("Все товары")
        _write_items_sheet(ws_items, receipts, db)

    # --- Листы по продавцам ---
    if group_by_seller:
        for info in sorted_groups:
            gk = info["inn"] if info["inn"] else _seller_group_key(info["display_name"])
            sheet_name = _unique_sheet_name(wb, info["display_name"])
            group_to_sheet[gk] = sheet_name

            ws_seller = wb.create_sheet(sheet_name)

            seller_receipts = []
            for r in receipts:
                r_inn = (r.get("seller_inn") or "").strip()
                r_name = _normalize_seller_name(r.get("seller_name", ""))
                r_gk = r_inn if r_inn else _seller_group_key(r_name)
                if r_gk == gk:
                    seller_receipts.append(r)

            ws_seller.merge_cells("A1:G1")
            ws_seller["A1"].value = _abbreviate_org_form(info["display_name"])
            ws_seller["A1"].font = SELLER_FONT
            ws_seller["A1"].fill = SELLER_FILL

            seller_total = info["total"]
            places = info["places"]
            places_text = f" ({', '.join(sorted(places))})" if places else ""
            ws_seller["A2"] = (
                f"ИНН: {info['inn']} | Чеков: {info['count']}"
                f" | Сумма: {seller_total:,.2f} ₽{places_text}"
            )
            ws_seller["A2"].font = Font(bold=True)

            if info["addresses"]:
                addrs = "; ".join(sorted(info["addresses"]))
                ws_seller["A3"] = f"Адреса: {addrs}"
                ws_seller["A3"].font = Font(size=9, italic=True, color="666666")
                items_start = 5
            else:
                items_start = 4

            if include_items and db:
                _write_receipts_with_items(
                    ws_seller, seller_receipts, db, start_row=items_start
                )
            else:
                _write_receipts_sheet(ws_seller, seller_receipts, start_row=items_start)

    # --- Заполняем гиперссылки в Сводке ---
    for link_row, gk in summary_link_rows:
        sheet_name = group_to_sheet.get(gk, "")
        if sheet_name:
            cell = ws_summary.cell(row=link_row, column=6)
            cell.value = "→ перейти"
            cell.font = LINK_FONT
            cell.hyperlink = f"#{quote_sheetname(sheet_name)}!A1"

    _auto_width(ws_summary)

    # Сохраняем
    wb.save(output_path)

    # Восстанавливаем стили диаграмм (тёмная тема, цвета), которые openpyxl
    # не сохраняет при re-save шаблона.
    if os.path.exists(FOOD_TEMPLATE_PATH):
        try:
            _inject_chart_styles(output_path, FOOD_TEMPLATE_PATH)
        except Exception as e:
            logger.warning("Не удалось восстановить стили диаграмм: %s", e)

    logger.info("Excel файл сохранён: %s", output_path)
    return output_path


def _inject_chart_styles(output_path: str, template_path: str):
    """Копировать файлы style*.xml/colors*.xml из шаблона в выходной файл
    и дополнить chart*.xml.rels ссылками на них.

    openpyxl удаляет эти файлы при re-save, поэтому восстанавливаем вручную.
    """
    import zipfile, shutil, re

    # 1) Читаем style/colors XML из шаблона
    style_files: dict[str, bytes] = {}
    with zipfile.ZipFile(template_path) as zt:
        for name in zt.namelist():
            if name.startswith('xl/charts/') and name.endswith('.xml'):
                base = name.split('/')[-1]
                if base.startswith('style') or base.startswith('colors'):
                    style_files[name] = zt.read(name)
    if not style_files:
        return

    # 2) Определяем сколько диаграмм в выходном файле
    tmp_path = output_path + '.tmp'
    shutil.move(output_path, tmp_path)

    # Счёт chart*.xml в выходе
    with zipfile.ZipFile(tmp_path) as zin:
        out_charts = sorted([
            int(re.search(r'chart(\d+)\.xml$', n).group(1))
            for n in zin.namelist()
            if re.search(r'xl/charts/chart\d+\.xml$', n)
        ])

    # Маппинг: chart{N} → style{M}/colors{M} (из шаблона)
    # В шаблоне индексы 1..5. В выходе столько же (5 диаграмм). Мапим 1-к-1.
    tpl_indexes = sorted([
        int(re.search(r'style(\d+)\.xml$', k).group(1))
        for k in style_files if 'style' in k.split('/')[-1]
    ])

    with zipfile.ZipFile(tmp_path) as zin, \
         zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:

        existing_names = set(zin.namelist())

        for item in zin.namelist():
            data = zin.read(item)

            # Патчим существующие chart{N}.xml.rels
            m = re.match(r'xl/charts/_rels/chart(\d+)\.xml\.rels$', item)
            if m:
                n = int(m.group(1))
                if n <= len(tpl_indexes):
                    tpl_n = tpl_indexes[n - 1]
                    ds = data.decode('utf-8')
                    if 'chartStyle' not in ds:
                        ds = ds.replace(
                            '</Relationships>',
                            f'<Relationship Id="rIdStyle{n}" '
                            f'Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" '
                            f'Target="style{tpl_n}.xml"/>'
                            f'<Relationship Id="rIdColor{n}" '
                            f'Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" '
                            f'Target="colors{tpl_n}.xml"/>'
                            '</Relationships>'
                        )
                        data = ds.encode('utf-8')

            # Патчим [Content_Types].xml
            if item == '[Content_Types].xml':
                ds = data.decode('utf-8')
                overrides = []
                for sf_path in sorted(style_files.keys()):
                    base = sf_path.split('/')[-1]
                    if base.startswith('style'):
                        ct = 'application/vnd.ms-office.chartstyle+xml'
                    else:
                        ct = 'application/vnd.ms-office.chartcolorstyle+xml'
                    part_name = '/' + sf_path
                    if f'PartName="{part_name}"' not in ds:
                        overrides.append(
                            f'<Override PartName="{part_name}" ContentType="{ct}"/>'
                        )
                if overrides:
                    ds = ds.replace('</Types>', ''.join(overrides) + '</Types>')
                    data = ds.encode('utf-8')

            zout.writestr(item, data)

        # Создаём chart{N}.xml.rels, если их нет
        for n in out_charts:
            rels_path = f'xl/charts/_rels/chart{n}.xml.rels'
            if rels_path not in existing_names and n <= len(tpl_indexes):
                tpl_n = tpl_indexes[n - 1]
                rels_xml = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="rIdStyle{n}" '
                    f'Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" '
                    f'Target="style{tpl_n}.xml"/>'
                    f'<Relationship Id="rIdColor{n}" '
                    f'Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" '
                    f'Target="colors{tpl_n}.xml"/>'
                    '</Relationships>'
                )
                zout.writestr(rels_path, rels_xml)

        # Добавляем style/colors файлы
        for name, content in style_files.items():
            if name not in existing_names:
                zout.writestr(name, content)

    import os as _os
    _os.unlink(tmp_path)


def _write_receipts_sheet(ws, receipts: list[dict], start_row: int = 1):
    """Записать чеки на лист."""
    row = start_row

    for col, header in enumerate(RECEIPT_HEADERS, 1):
        ws.cell(row=row, column=col, value=header)
    _apply_header_style(ws, row, len(RECEIPT_HEADERS))

    for receipt in receipts:
        row += 1
        date_val = _format_date(receipt.get("receipt_date", ""))

        ws.cell(row=row, column=1, value=date_val)
        ws.cell(row=row, column=2, value=receipt.get("fiscal_document_id", ""))
        ws.cell(row=row, column=3, value=_store_name(receipt))

        sum_cell = ws.cell(row=row, column=4, value=receipt.get("total_sum", 0))
        sum_cell.number_format = MONEY_FORMAT

        ws.cell(row=row, column=5, value=receipt.get("seller_address", ""))
        ws.cell(row=row, column=6, value=receipt.get("seller_inn", ""))

        for col in range(1, len(RECEIPT_HEADERS) + 1):
            ws.cell(row=row, column=col).border = THIN_BORDER

    _auto_width(ws)


def _write_receipts_with_items(ws, receipts: list[dict], db, start_row: int = 1):
    """Записать чеки с развёрнутыми товарами на лист продавца."""
    row = start_row

    headers = ["Дата", "Номер ФД", "Торговая точка", "Наименование товара",
               "Цена", "Кол-во", "Сумма"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=row, column=col, value=header)
    _apply_header_style(ws, row, len(headers))

    for receipt in receipts:
        receipt_id = receipt.get("id")
        date_val = _format_date(receipt.get("receipt_date", ""))
        store = _store_name(receipt)
        fd_id = receipt.get("fiscal_document_id", "")
        total = receipt.get("total_sum", 0)

        items = db.get_receipt_items(receipt_id) if receipt_id else []

        if items:
            # Строка-заголовок чека
            row += 1
            ws.cell(row=row, column=1, value=date_val).font = RECEIPT_HEADER_FONT
            ws.cell(row=row, column=2, value=fd_id).font = RECEIPT_HEADER_FONT
            ws.cell(row=row, column=3, value=store).font = RECEIPT_HEADER_FONT

            total_cell = ws.cell(row=row, column=7, value=total)
            total_cell.number_format = MONEY_FORMAT
            total_cell.font = RECEIPT_HEADER_FONT

            for col in range(1, len(headers) + 1):
                ws.cell(row=row, column=col).fill = RECEIPT_HEADER_FILL
                ws.cell(row=row, column=col).border = THIN_BORDER

            # Позиции чека
            for item in items:
                row += 1
                ws.cell(row=row, column=4, value=item.get("name", ""))

                price_cell = ws.cell(row=row, column=5, value=item.get("price", 0))
                price_cell.number_format = MONEY_FORMAT

                ws.cell(row=row, column=6, value=item.get("quantity", 0))

                item_total_cell = ws.cell(row=row, column=7, value=item.get("total", 0))
                item_total_cell.number_format = MONEY_FORMAT

                for col in range(1, len(headers) + 1):
                    ws.cell(row=row, column=col).border = THIN_BORDER
        else:
            # Чек без товаров — одна строка
            row += 1
            ws.cell(row=row, column=1, value=date_val)
            ws.cell(row=row, column=2, value=fd_id)
            ws.cell(row=row, column=3, value=store)

            total_cell = ws.cell(row=row, column=7, value=total)
            total_cell.number_format = MONEY_FORMAT

            for col in range(1, len(headers) + 1):
                ws.cell(row=row, column=col).border = THIN_BORDER

    _auto_width(ws)


def _write_items_sheet(ws, receipts: list[dict], db):
    """Записать все позиции чеков на лист."""
    row = 1
    for col, header in enumerate(ITEM_HEADERS, 1):
        ws.cell(row=row, column=col, value=header)
    _apply_header_style(ws, row, len(ITEM_HEADERS))

    for receipt in receipts:
        receipt_id = receipt.get("id")
        if not receipt_id:
            continue

        items = db.get_receipt_items(receipt_id)
        date_val = _format_date(receipt.get("receipt_date", ""))
        seller = receipt.get("seller_name", "")
        store = _store_name(receipt)

        for item in items:
            row += 1
            ws.cell(row=row, column=1, value=date_val)
            ws.cell(row=row, column=2, value=seller)
            ws.cell(row=row, column=3, value=store)
            ws.cell(row=row, column=4, value=item.get("name", ""))

            price_cell = ws.cell(row=row, column=5, value=item.get("price", 0))
            price_cell.number_format = MONEY_FORMAT

            ws.cell(row=row, column=6, value=item.get("quantity", 0))

            total_cell = ws.cell(row=row, column=7, value=item.get("total", 0))
            total_cell.number_format = MONEY_FORMAT

            for col in range(1, len(ITEM_HEADERS) + 1):
                ws.cell(row=row, column=col).border = THIN_BORDER

    _auto_width(ws)


def _parse_date(date_str) -> Optional[datetime]:
    """Распарсить дату из строки."""
    try:
        s = str(date_str)
        if "T" in s:
            return datetime.fromisoformat(s.replace("Z", "+00:00"))
        return datetime.strptime(s[:10], "%Y-%m-%d")
    except Exception:
        return None


def _style_chart(chart, hide_legend: bool = True):
    """Настроить подписи значений на столбцах и убрать боковую легенду."""
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.numFmt = '#,##0'
    if hide_legend:
        chart.legend = None


MONTH_NAMES_RU = [
    "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]


def _chart_rows(chart_height_cm: float) -> int:
    """Примерное количество строк, которое занимает график данной высоты."""
    return max(int(chart_height_cm / 0.75) + 2, 18)


def _clear_table_rows(ws, header_row: int, max_scan: int = 200):
    """Очистить A-B начиная со строки (header_row+1) до первой пустой строки.

    Используется перед записью свежих данных в таблицу шаблона.
    """
    r = header_row + 1
    cleared = 0
    while cleared < max_scan:
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        if a is None and b is None:
            break
        ws.cell(row=r, column=1).value = None
        ws.cell(row=r, column=2).value = None
        # Сбросить границы ячеек (убрать оформление от предыдущего экспорта)
        ws.cell(row=r, column=1).border = Border()
        ws.cell(row=r, column=2).border = Border()
        r += 1
        cleared += 1
    return r - 1  # последняя очищенная строка


def _write_food_data_sheet(ws, food_data: list[dict]):
    """Лист-расшифровка: все позиции, вошедшие в дашборд «Продукты».

    Колонки: Дата, Магазин, Категория, Наименование, Цена, Кол-во, Сумма,
             Продавец, ИНН продавца.
    Позволяет пользователю сверить итоговые суммы и найти расхождения
    с собственным учётом.
    """
    headers = [
        "Дата", "Магазин", "Категория", "Наименование",
        "Цена", "Кол-во", "Сумма", "Продавец", "ИНН продавца",
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    _apply_header_style(ws, 1, len(headers))

    row = 1
    for d in sorted(food_data, key=lambda x: str(x.get("receipt_date", ""))):
        row += 1
        ws.cell(row=row, column=1, value=_format_date(d.get("receipt_date", "")))
        ws.cell(row=row, column=2, value=d.get("store", ""))
        ws.cell(row=row, column=3, value=d.get("category", ""))
        ws.cell(row=row, column=4, value=d.get("name", ""))

        price_cell = ws.cell(row=row, column=5, value=d.get("price", 0))
        price_cell.number_format = MONEY_FORMAT

        ws.cell(row=row, column=6, value=d.get("quantity", 0))

        total_cell = ws.cell(row=row, column=7, value=d.get("total", 0))
        total_cell.number_format = MONEY_FORMAT

        ws.cell(row=row, column=8, value=d.get("seller_name", ""))
        ws.cell(row=row, column=9, value=d.get("seller_inn", ""))

        for col in range(1, len(headers) + 1):
            ws.cell(row=row, column=col).border = THIN_BORDER

    # Итого
    row += 1
    ws.cell(row=row, column=1, value="ИТОГО").font = TOTAL_FONT
    ws.cell(row=row, column=6, value=f"Строк: {len(food_data)}").font = TOTAL_FONT
    total_sum = sum(d.get("total", 0) or 0 for d in food_data)
    total_cell = ws.cell(row=row, column=7, value=total_sum)
    total_cell.font = TOTAL_FONT
    total_cell.number_format = MONEY_FORMAT
    for col in range(1, len(headers) + 1):
        ws.cell(row=row, column=col).fill = TOTAL_FILL
        ws.cell(row=row, column=col).border = THIN_BORDER

    _auto_width(ws)


def _write_category_sheet(ws, category_name: str, items: list[dict]):
    """Лист категории расходов. Аналог 'Продукты_данные', но с мини-сводкой
    сверху: ежемесячные траты, траты по магазинам и общий итог.

    Колонки:
        A  Дата
        B  Магазин
        C  Наименование
        D  Цена
        E  Кол-во
        F  Сумма
        G  Продавец
        H  ИНН продавца
        I  Номер ФД
    """
    from collections import defaultdict as dd

    # ── Заголовок листа ──
    ws.merge_cells("A1:F1")
    ws["A1"].value = f"Категория: {category_name}"
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color="2B5797")
    ws["A1"].alignment = Alignment(horizontal="left")

    total_sum = sum(it.get("total", 0) or 0 for it in items)
    ws["A2"] = f"Всего позиций: {len(items)}"
    ws["C2"] = f"Общая сумма: {total_sum:,.2f} ₽"
    ws["C2"].font = Font(bold=True)

    # ── Агрегации ──
    monthly = dd(float)
    store_totals = dd(float)
    yearly = dd(float)
    for it in items:
        dt = _parse_date(it.get("receipt_date"))
        if dt:
            monthly[f"{dt.year}-{dt.month:02d}"] += it.get("total", 0) or 0
            yearly[dt.year] += it.get("total", 0) or 0
        store_totals[it.get("store") or "Неизвестно"] += it.get("total", 0) or 0

    sorted_months = sorted(monthly.items())
    sorted_years = sorted(yearly.items())
    sorted_stores = sorted(store_totals.items(), key=lambda x: x[1], reverse=True)

    GAP = 2

    def _mini_section(start_row: int, title: str, header_labels: list[str],
                      rows_data: list[tuple]) -> int:
        ws.cell(row=start_row, column=1, value=title).font = Font(
            name="Calibri", bold=True, size=12, color="2B5797")
        hdr = start_row + 1
        for ci, h in enumerate(header_labels):
            ws.cell(row=hdr, column=ci + 1, value=h)
        _apply_header_style(ws, hdr, len(header_labels))
        r = hdr
        for item_row in rows_data:
            r += 1
            ws.cell(row=r, column=1, value=item_row[0])
            c2 = ws.cell(row=r, column=2, value=item_row[1])
            c2.number_format = MONEY_FORMAT
            for ci in range(1, len(header_labels) + 1):
                ws.cell(row=r, column=ci).border = THIN_BORDER
        return r

    cursor = 4
    months_items = [(f"{MONTH_NAMES_RU[int(m.split('-')[1])]} {m.split('-')[0]}",
                     round(t, 2)) for m, t in sorted_months]
    cursor = _mini_section(cursor, "Ежемесячные траты",
                           ["Месяц", "Сумма, ₽"], months_items) + GAP + 1

    years_items = [(str(y), round(t, 2)) for y, t in sorted_years]
    cursor = _mini_section(cursor, "Траты по годам",
                           ["Год", "Сумма, ₽"], years_items) + GAP + 1

    stores_items = [(s, round(t, 2)) for s, t in sorted_stores]
    cursor = _mini_section(cursor, "Траты по магазинам",
                           ["Магазин", "Сумма, ₽"], stores_items) + GAP + 1

    # ── Таблица всех позиций ──
    ws.cell(row=cursor, column=1, value="Все позиции").font = Font(
        name="Calibri", bold=True, size=12, color="2B5797")
    cursor += 1

    headers = ["Дата", "Магазин", "Наименование", "Цена", "Кол-во",
               "Сумма", "Продавец", "ИНН продавца", "Номер ФД"]
    for ci, h in enumerate(headers):
        ws.cell(row=cursor, column=ci + 1, value=h)
    _apply_header_style(ws, cursor, len(headers))

    row = cursor
    for it in sorted(items, key=lambda x: str(x.get("receipt_date", ""))):
        row += 1
        ws.cell(row=row, column=1, value=_format_date(it.get("receipt_date", "")))
        ws.cell(row=row, column=2, value=it.get("store", ""))
        ws.cell(row=row, column=3, value=it.get("name", ""))

        price_cell = ws.cell(row=row, column=4, value=it.get("price", 0))
        price_cell.number_format = MONEY_FORMAT

        ws.cell(row=row, column=5, value=it.get("quantity", 0))

        total_cell = ws.cell(row=row, column=6, value=it.get("total", 0))
        total_cell.number_format = MONEY_FORMAT

        ws.cell(row=row, column=7, value=it.get("seller_name", ""))
        ws.cell(row=row, column=8, value=it.get("seller_inn", ""))
        ws.cell(row=row, column=9, value=it.get("fiscal_document_id", ""))

        for col in range(1, len(headers) + 1):
            ws.cell(row=row, column=col).border = THIN_BORDER

    # ИТОГО
    row += 1
    ws.cell(row=row, column=1, value="ИТОГО").font = TOTAL_FONT
    total_cell = ws.cell(row=row, column=6, value=total_sum)
    total_cell.font = TOTAL_FONT
    total_cell.number_format = MONEY_FORMAT
    for col in range(1, len(headers) + 1):
        ws.cell(row=row, column=col).fill = TOTAL_FILL
        ws.cell(row=row, column=col).border = THIN_BORDER

    # Ширины столбцов (фиксированные, чтоб мини-таблицы не растягивались)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["H"].width = 16
    ws.column_dimensions["I"].width = 12


def _write_statistics_sheet(ws, by_category: dict[str, list[dict]]):
    """Лист 'Статистика' — матрица категорий × месяцев/лет.

    Формат:
        Заголовок
        [Таблица 1] Ежемесячные траты по категориям
        [Таблица 2] Годовые траты по категориям
    """
    from collections import defaultdict as dd

    ws.merge_cells("A1:F1")
    ws["A1"].value = "Статистика расходов по категориям"
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color="2B5797")
    ws["A1"].alignment = Alignment(horizontal="left")

    # Собираем все уникальные категории, месяцы и годы
    all_cats = list(by_category.keys())
    # Сортируем так, чтобы "Без категории" была в конце
    def _cat_sort_key(c: str) -> tuple:
        return (1 if c == "Без категории" else 0, c)
    all_cats.sort(key=_cat_sort_key)

    monthly: dict[str, dict[str, float]] = dd(lambda: dd(float))
    yearly: dict[int, dict[str, float]] = dd(lambda: dd(float))
    all_months: set = set()
    all_years: set = set()
    grand_total: dict[str, float] = dd(float)

    for cat, items in by_category.items():
        for it in items:
            dt = _parse_date(it.get("receipt_date"))
            if not dt:
                continue
            mkey = f"{dt.year}-{dt.month:02d}"
            t = it.get("total", 0) or 0
            monthly[mkey][cat] += t
            yearly[dt.year][cat] += t
            grand_total[cat] += t
            all_months.add(mkey)
            all_years.add(dt.year)

    sorted_months = sorted(all_months, reverse=True)  # Новые сверху
    sorted_years = sorted(all_years, reverse=True)

    # ── Таблица 1: ежемесячные траты ──
    row = 3
    ws.cell(row=row, column=1, value="Ежемесячные траты по категориям").font = Font(
        name="Calibri", bold=True, size=12, color="2B5797")
    row += 1

    headers = ["Месяц"] + all_cats + ["Итого"]
    for ci, h in enumerate(headers):
        ws.cell(row=row, column=ci + 1, value=h)
    _apply_header_style(ws, row, len(headers))

    for mkey in sorted_months:
        row += 1
        y, m = mkey.split("-")
        ws.cell(row=row, column=1, value=f"{MONTH_NAMES_RU[int(m)]} {y}")
        month_total = 0.0
        for ci, cat in enumerate(all_cats):
            v = round(monthly[mkey].get(cat, 0), 2)
            c = ws.cell(row=row, column=ci + 2, value=v if v else None)
            if v:
                c.number_format = MONEY_FORMAT
            month_total += v
        tc = ws.cell(row=row, column=len(headers),
                     value=round(month_total, 2))
        tc.number_format = MONEY_FORMAT
        tc.font = Font(bold=True)
        for ci in range(1, len(headers) + 1):
            ws.cell(row=row, column=ci).border = THIN_BORDER

    # Итог по месяцам (сумма по всем категориям)
    row += 1
    ws.cell(row=row, column=1, value="ИТОГО").font = TOTAL_FONT
    total_all = 0.0
    for ci, cat in enumerate(all_cats):
        v = round(grand_total.get(cat, 0), 2)
        c = ws.cell(row=row, column=ci + 2, value=v)
        c.number_format = MONEY_FORMAT
        c.font = TOTAL_FONT
        total_all += v
    tc = ws.cell(row=row, column=len(headers), value=round(total_all, 2))
    tc.number_format = MONEY_FORMAT
    tc.font = TOTAL_FONT
    for ci in range(1, len(headers) + 1):
        ws.cell(row=row, column=ci).fill = TOTAL_FILL
        ws.cell(row=row, column=ci).border = THIN_BORDER

    # ── Таблица 2: годовые траты ──
    row += 3
    ws.cell(row=row, column=1, value="Годовые траты по категориям").font = Font(
        name="Calibri", bold=True, size=12, color="2B5797")
    row += 1

    for ci, h in enumerate(headers):
        ws.cell(row=row, column=ci + 1, value=h if ci != 0 else "Год")
    _apply_header_style(ws, row, len(headers))

    for y in sorted_years:
        row += 1
        ws.cell(row=row, column=1, value=str(y))
        year_total = 0.0
        for ci, cat in enumerate(all_cats):
            v = round(yearly[y].get(cat, 0), 2)
            c = ws.cell(row=row, column=ci + 2, value=v if v else None)
            if v:
                c.number_format = MONEY_FORMAT
            year_total += v
        tc = ws.cell(row=row, column=len(headers),
                     value=round(year_total, 2))
        tc.number_format = MONEY_FORMAT
        tc.font = Font(bold=True)
        for ci in range(1, len(headers) + 1):
            ws.cell(row=row, column=ci).border = THIN_BORDER

    # Ширина столбцов
    ws.column_dimensions["A"].width = 18
    for ci in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 16


def _write_food_dashboard(wb, food_data: list[dict]):
    """Обновить лист-дашборд 'Продукты' данными из food_data.

    Лист 'Продукты' должен уже существовать в wb (загружен из шаблона
    food_template.xlsx). Функция полностью очищает данные ниже строки 3
    и заново записывает 5 таблиц с динамическими позициями
    (2 пустые строки между таблицами).

    Порядок таблиц:
      1) Ежемесячные траты на продукты
      2) Траты на продукты по годам
      3) Траты по магазинам (всё время)
      4) Траты по категориям продуктов (ежемесячно) — матрица
      5) Итого по категориям за всё время

    Диаграммы берутся из шаблона и привязываются к новым диапазонам.
    """
    from collections import defaultdict as dd

    if "Продукты" not in wb.sheetnames:
        logger.warning("Шаблон 'Продукты' не найден в wb, лист пропущен")
        return
    ws = wb["Продукты"]

    # ══════════════════════════════════════════════════════════
    # 0. Идентифицируем диаграммы шаблона ДО очистки данных
    # ══════════════════════════════════════════════════════════
    # Шаблон:
    #   Chart with $B$6   → магазины (stores)
    #   Chart with $B$35  → месяцы (months)
    #   Chart with $B$71  → годы (years, PieChart)
    #   Chart with $B$96  → категории по месяцам (cat_monthly matrix)
    #   Chart with $B$135 → итого по категориям (cat_totals)
    chart_months = None
    chart_years = None
    chart_stores = None
    chart_cat_monthly = None
    chart_cat_totals = None

    for chart in ws._charts:
        if not chart.ser or not chart.ser[0].val or not chart.ser[0].val.numRef:
            continue
        ref = chart.ser[0].val.numRef.f or ""
        if "$B$6" in ref:
            chart_stores = chart
        elif "$B$35" in ref:
            chart_months = chart
        elif "$B$71" in ref:
            chart_years = chart
        elif "$B$96" in ref or "$B$95" in ref:
            chart_cat_monthly = chart
        elif "$B$135" in ref or "$B$134" in ref:
            chart_cat_totals = chart

    # ══════════════════════════════════════════════════════════
    # 1. Очистить ВСЕ данные ниже строки 3
    # ══════════════════════════════════════════════════════════
    # Определяем макс. использованный столбец (матрица категорий может быть широкой)
    max_used_col = ws.max_column or 2
    for r in range(4, (ws.max_row or 200) + 1):
        for c in range(1, max_used_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.border = Border()
            cell.font = Font()
            cell.fill = PatternFill()
            cell.number_format = 'General'

    # Обновим верхний колонтитул
    ws["A1"].value = "Дашборд — Продукты питания"
    ws["A2"].value = f"Всего позиций: {len(food_data)}"
    total_food = sum(d.get("total", 0) or 0 for d in food_data)
    ws["C2"].value = f"Общая сумма: {total_food:,.2f} ₽"

    # ══════════════════════════════════════════════════════════
    # 2. Агрегируем данные
    # ══════════════════════════════════════════════════════════
    store_totals = dd(float)
    monthly = dd(float)
    yearly = dd(float)
    cat_totals = dd(float)
    for d in food_data:
        store_totals[d.get("store") or "Неизвестно"] += d.get("total", 0) or 0
        dt = _parse_date(d.get("receipt_date"))
        if dt:
            monthly[f"{dt.year}-{dt.month:02d}"] += d.get("total", 0) or 0
            yearly[dt.year] += d.get("total", 0) or 0
        cat_totals[d.get("category", "Без категории")] += d.get("total", 0) or 0

    sorted_stores = sorted(store_totals.items(), key=lambda x: x[1], reverse=True)
    sorted_months = sorted(monthly.items())
    sorted_years = sorted(yearly.items())
    sorted_cat_totals = sorted(cat_totals.items(),
                                key=lambda x: x[1], reverse=True)

    cat_monthly = dd(lambda: dd(float))
    all_cats = set()
    all_months_set = set()
    for d in food_data:
        dt = _parse_date(d.get("receipt_date"))
        if dt:
            mkey = f"{dt.year}-{dt.month:02d}"
            cat = d.get("category", "Без категории")
            cat_monthly[mkey][cat] += d.get("total", 0) or 0
            all_cats.add(cat)
            all_months_set.add(mkey)
    sorted_cats = sorted(all_cats)
    sorted_cat_months = sorted(all_months_set)

    # Подготовка данных таблиц
    months_items = [(f"{MONTH_NAMES_RU[int(m.split('-')[1])]} {m.split('-')[0]}",
                     round(t, 2)) for m, t in sorted_months]
    years_items = [(str(y), round(t, 2)) for y, t in sorted_years]
    stores_items = [(s, round(t, 2)) for s, t in sorted_stores]
    cat_total_items = [(c, round(t, 2)) for c, t in sorted_cat_totals]

    # ══════════════════════════════════════════════════════════
    # 3. Записываем таблицы динамически (2 пустые строки между ними)
    # ══════════════════════════════════════════════════════════
    GAP = 2  # пустых строк между таблицами

    def _write_section(start_row: int, title: str, header_labels: list[str],
                       items: list[tuple], col_count: int = 2):
        """Записать секцию: заголовок, header, данные. Возвращает (header_row, data_end)."""
        # Строка-заголовок секции
        ws.cell(row=start_row, column=1, value=title).font = Font(
            name="Calibri", bold=True, size=12, color="2B5797")
        # Строка-заголовок таблицы
        hdr_row = start_row + 1
        for ci, h in enumerate(header_labels):
            ws.cell(row=hdr_row, column=ci + 1, value=h)
        _apply_header_style(ws, hdr_row, col_count)
        # Данные
        r = hdr_row
        for item_row in items:
            r += 1
            if isinstance(item_row, tuple) and len(item_row) == 2:
                ws.cell(row=r, column=1, value=item_row[0])
                c = ws.cell(row=r, column=2, value=item_row[1])
                c.number_format = MONEY_FORMAT
            else:
                # Для матрицы — item_row это list значений
                for ci, val in enumerate(item_row):
                    cell = ws.cell(row=r, column=ci + 1, value=val)
                    if ci >= 1:  # числовые столбцы
                        cell.number_format = MONEY_FORMAT
            for ci in range(1, col_count + 1):
                ws.cell(row=r, column=ci).border = THIN_BORDER
        data_end = r
        return hdr_row, data_end

    cursor = 4  # начинаем с 4-й строки (1-3: заголовки дашборда)

    # 1) Ежемесячные траты на продукты
    months_hdr, months_end = _write_section(
        cursor, "Ежемесячные траты на продукты",
        ["Месяц", "Сумма, ₽"], months_items)

    # 2) Траты на продукты по годам
    cursor = months_end + GAP + 1
    years_hdr, years_end = _write_section(
        cursor, "Траты на продукты по годам",
        ["Год", "Сумма, ₽"], years_items)

    # 3) Траты по магазинам (всё время)
    cursor = years_end + GAP + 1
    stores_hdr, stores_end = _write_section(
        cursor, "Траты по магазинам (всё время)",
        ["Магазин", "Сумма, ₽"], stores_items)

    # 4) Траты по категориям продуктов (ежемесячно) — матрица
    cursor = stores_end + GAP + 1
    cat_m_col_count = 1 + len(sorted_cats)
    cat_m_header_labels = ["Месяц"] + sorted_cats
    cat_m_rows = []
    for mkey in sorted_cat_months:
        y, m = mkey.split("-")
        row_vals = [f"{MONTH_NAMES_RU[int(m)]} {y}"]
        for cat in sorted_cats:
            row_vals.append(round(cat_monthly[mkey].get(cat, 0), 2))
        cat_m_rows.append(row_vals)
    cat_m_hdr, cat_m_end = _write_section(
        cursor, "Траты по категориям продуктов (ежемесячно)",
        cat_m_header_labels, cat_m_rows, col_count=cat_m_col_count)

    # 5) Итого по категориям за всё время
    cursor = cat_m_end + GAP + 1
    cat_t_hdr, cat_t_end = _write_section(
        cursor, "Итого по категориям за всё время",
        ["Категория", "Сумма, ₽"], cat_total_items)

    # ══════════════════════════════════════════════════════════
    # 4. Обновление диапазонов диаграмм
    # ══════════════════════════════════════════════════════════
    def _update_chart_range(chart, header_row: int, data_end: int,
                            min_col: int = 2, max_col: int = 2,
                            tx_row: Optional[int] = None,
                            cat_first: Optional[int] = None,
                            cat_last: Optional[int] = None,
                            cat_col: int = 1):
        """Обновить ссылки существующей диаграммы шаблона на новый диапазон."""
        sheet_title = ws.title
        from openpyxl.chart.data_source import NumDataSource, NumRef, StrRef, AxDataSource
        from openpyxl.chart.series import SeriesLabel, Series
        sn = quote_sheetname(sheet_title)
        if not chart.ser:
            return
        chart.ser.clear()
        n_series = max_col - min_col + 1
        for si in range(n_series):
            col = min_col + si
            col_letter = get_column_letter(col)
            val_f = f"{sn}!${col_letter}${header_row + 1}:${col_letter}${data_end}"
            ser = Series()
            ser.val = NumDataSource(numRef=NumRef(f=val_f))
            if tx_row is not None:
                tx_f = f"{sn}!${col_letter}${tx_row}"
                ser.tx = SeriesLabel(strRef=StrRef(f=tx_f))
            if cat_first is not None and cat_last is not None:
                cc = get_column_letter(cat_col)
                cat_f = f"{sn}!${cc}${cat_first}:${cc}${cat_last}"
                ser.cat = AxDataSource(strRef=StrRef(f=cat_f))
            ser.idx = si
            ser.order = si
            chart.ser.append(ser)

    # Месяцы → chart_months
    if chart_months and months_end >= months_hdr + 1:
        _update_chart_range(
            chart_months, months_hdr, months_end,
            cat_first=months_hdr + 1, cat_last=months_end)

    # Годы → chart_years
    if chart_years and years_end >= years_hdr + 1:
        _update_chart_range(
            chart_years, years_hdr, years_end,
            cat_first=years_hdr + 1, cat_last=years_end)

    # Магазины → chart_stores
    if chart_stores and stores_end >= stores_hdr + 1:
        _update_chart_range(
            chart_stores, stores_hdr, stores_end,
            cat_first=stores_hdr + 1, cat_last=stores_end)

    # Категории по месяцам → chart_cat_monthly (матрица)
    if chart_cat_monthly and sorted_cats and sorted_cat_months and cat_m_end >= cat_m_hdr + 1:
        _update_chart_range(
            chart_cat_monthly, cat_m_hdr, cat_m_end,
            min_col=2, max_col=1 + len(sorted_cats),
            tx_row=cat_m_hdr,
            cat_first=cat_m_hdr + 1, cat_last=cat_m_end)

    # Итого по категориям → chart_cat_totals
    if chart_cat_totals and cat_t_end >= cat_t_hdr + 1:
        _update_chart_range(
            chart_cat_totals, cat_t_hdr, cat_t_end,
            tx_row=cat_t_hdr,
            cat_first=cat_t_hdr + 1, cat_last=cat_t_end)

    # Ширины столбцов A-B задаём вручную, остальное — из шаблона
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 16
