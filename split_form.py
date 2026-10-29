# -*- coding: utf-8 -*-
"""
split_form.py
Разбор файлов «Сплит-форма» (сметные цены ФГИС ЦС) и выгрузка справочника
цен в Excel.

Правило расчёта цены (без НДС):
    1. Если 8-я колонка («Сметная цена в текущем уровне цен») заполнена
       числом — берётся она.
    2. Иначе цена = 5-я колонка («Сметная цена на 01.01.2022») ×
       9-я колонка («Индекс изменения сметной стоимости»).
    3. Цена с НДС 22% = цена без НДС × 1,22.

В справочник не попадают тарифы на перевозку (ед. изм. «1т груза») и
оплата труда рабочих (ед. изм. «чел.-ч»). Машины и механизмы определяются
по коду ресурса «91.*» или единице измерения «маш.-ч».
"""
import math
import re

from python_calamine import CalamineWorkbook

# Ставка НДС: цена с НДС = цена без НДС × 1,22
VAT_KOEFF = 1.22

KIND_MATERIAL = "материал"
KIND_MACHINE = "машины и механизмы"

# Единицы измерения, которые не попадают в справочник
SKIP_UNITS = ("1т груза", "чел.-ч")

# Позиции колонок в сплит-форме (используются, если шапку не удалось разобрать)
DEFAULT_POSITIONS = {
    "code": 0,
    "name": 1,
    "unit": 2,
    "price_otpusk": 3,
    "price_2022": 4,
    "group_code": 5,
    "group_name": 6,
    "price_current": 7,
    "index": 8,
}

HEADER_SCAN_ROWS = 200
PROGRESS_STEP = 2000


# --------------------------------------------------------------------------
# Вспомогательные функции
# --------------------------------------------------------------------------
def _norm(value) -> str:
    """Нормализует значение заголовка: убирает переводы строк и лишние пробелы."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def _num(value):
    """Преобразует значение ячейки в float; для прочерков и текста возвращает None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    s = str(value).strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
    if not s or s in ("-", "—", "nan", "none"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _text(value) -> str:
    """Возвращает строковое значение ячейки без служебных токенов.

    Calamine отдаёт числа как float, поэтому целые значения приводятся к int,
    чтобы код группы выглядел как '511', а не '511.0'.

    Args:
        value: значение ячейки.

    Returns:
        str: строковое представление (пустая строка для None и NaN).
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return str(value)
    s = str(value).strip()
    return "" if s.lower() in ("nan", "none") else s


def _map_columns(header) -> dict:
    """Сопоставляет названия колонок шапки с индексами.

    Args:
        header: последовательность значений строки-шапки.

    Returns:
        dict: индексы колонок по ключам code, name, unit, price_current,
            price_otpusk, price_2022, index, group_code, group_name.
    """
    mapping = {}
    for idx, raw in enumerate(header):
        h = _norm(raw)
        if not h:
            continue
        if h.startswith("код ресурса"):
            mapping.setdefault("code", idx)
        elif h.startswith("наименование строительного ресурса") or h == "наименование":
            mapping.setdefault("name", idx)
        elif h.startswith("единица измерения"):
            mapping.setdefault("unit", idx)
        elif "сметная цена в текущем уровне" in h:
            mapping.setdefault("price_current", idx)
        elif h.startswith("отпускная цена"):
            mapping.setdefault("price_otpusk", idx)
        elif "сметная цена" in h and "01.01.2022" in h:
            mapping.setdefault("price_2022", idx)
        elif "индекс изменения" in h:
            mapping.setdefault("index", idx)
        elif h.startswith("номер группы"):
            mapping.setdefault("group_code", idx)
        elif h.startswith("наименование группы"):
            mapping.setdefault("group_name", idx)

    # Недостающие колонки добираем по стандартным позициям сплит-формы
    for key, pos in DEFAULT_POSITIONS.items():
        mapping.setdefault(key, pos)
    return mapping


def _find_header(rows_data: list):
    """Ищет строку-шапку сплит-формы по заголовку «Код ресурса».

    Args:
        rows_data: список строк (списков) листа Excel, полученный из Calamine.

    Returns:
        tuple: (номер строки шапки (1-based) или None, значения строки шапки или None).
    """
    limit = min(len(rows_data), HEADER_SCAN_ROWS)
    for row_idx in range(limit):
        row = rows_data[row_idx]
        for cell in row:
            if _norm(cell).startswith("код ресурса"):
                # Возвращаем номер строки как в openpyxl (1-based)
                return row_idx + 1, row
    return None, None


def _get(row, idx):
    """Безопасно возвращает ячейку строки по индексу."""
    if idx is None or idx < 0 or idx >= len(row):
        return None
    return row[idx]


def classify(code: str, unit: str) -> str:
    """Определяет тип позиции: машина/механизм или материал.

    Args:
        code (str): код ресурса из сплит-формы.
        unit (str): единица измерения.

    Returns:
        str: KIND_MACHINE для кодов «91.*» и единицы «маш.-ч», иначе KIND_MATERIAL.
    """
    if code.startswith("91.") or unit == "маш.-ч":
        return KIND_MACHINE
    return KIND_MATERIAL


# --------------------------------------------------------------------------
# Чтение сплит-формы
# --------------------------------------------------------------------------
def read_split_form(path: str, progress=None):
    """Читает файл сплит-формы и возвращает позиции для справочника цен.

    Цена без НДС берётся из 8-й колонки, а при её отсутствии вычисляется как
    произведение 5-й и 9-й колонок. Цена с НДС 22% считается умножением
    цены без НДС на 1,22. Тарифы на перевозку («1т груза») и оплата труда
    рабочих («чел.-ч») пропускаются.

    Args:
        path (str): путь к файлу .xlsx со сплит-формой.
        progress (callable, optional): функция прогресса
            ``progress(processed, total) -> bool``. Возврат False прерывает чтение.

    Returns:
        tuple[list[dict], dict]: список записей справочника и статистика чтения
            (total, transport, labor, no_price, imported, cancelled).
    """
    stats = {
        "total": 0,
        "transport": 0,
        "labor": 0,
        "no_price": 0,
        "imported": 0,
        "cancelled": False,
    }
    records = []

    # Загружаем книгу через Calamine (Rust-движок, работает быстро)
    wb = CalamineWorkbook.from_path(path)
    ws = _pick_sheet(wb)
    if ws is None:
        return records, stats

    # Получаем все данные листа в виде списка списков.
    # Это ключевое отличие: Calamine читает лист целиком, но делает это очень быстро.
    # skip_empty_area=False гарантирует, что мы не потеряем данные из-за пустых строк в начале.
    rows_data = ws.to_python(skip_empty_area=False)

    header_idx, header = _find_header(rows_data)
    if header_idx is None:
        # Шапка не найдена — используем стандартные позиции колонок
        cols = dict(DEFAULT_POSITIONS)
        data_start = 1
    else:
        cols = _map_columns(header)
        data_start = header_idx + 1

    # data_start — это номер строки (1-based), а срез списка — 0-based.
    # Поэтому индекс начала данных в rows_data: data_start - 1.
    start_index = data_start - 1
    total_rows = max(len(rows_data) - start_index, 1)
    data_rows = rows_data[start_index:]

    for processed, row in enumerate(data_rows, start=1):
        stats["total"] += 1

        code = _text(_get(row, cols["code"]))
        name = _text(_get(row, cols["name"]))
        unit = _text(_get(row, cols["unit"]))

        # Строка-нумератор колонок (1, 2, 3, ...) и пустые строки
        if not code or not name or (code.isdigit() and name.isdigit()):
            continue

        if unit in SKIP_UNITS:
            if unit == "1т груза":
                stats["transport"] += 1
            else:
                stats["labor"] += 1
            continue

        price_no_vat = _num(_get(row, cols["price_current"]))
        if price_no_vat is None:
            base = _num(_get(row, cols["price_2022"]))
            koeff = _num(_get(row, cols["index"]))
            if base is not None and koeff is not None:
                price_no_vat = base * koeff
        if price_no_vat is None:
            stats["no_price"] += 1
            continue

        price_no_vat = round(price_no_vat, 2)
        records.append({
            "code": code,
            "name": name,
            "unit": unit,
            "price_no_vat": price_no_vat,
            "price_vat": round(price_no_vat * VAT_KOEFF, 2),
            "kind": classify(code, unit),
            "group_code": _text(_get(row, cols["group_code"])),
            "group_name": _text(_get(row, cols["group_name"])),
        })
        stats["imported"] += 1

        if progress is not None and processed % PROGRESS_STEP == 0:
            if not progress(processed, total_rows):
                stats["cancelled"] = True
                break

    if progress is not None and not stats["cancelled"]:
        progress(total_rows, total_rows)

    return records, stats


def _pick_sheet(workbook: CalamineWorkbook):
    """Выбирает лист сплит-формы: с «сплит» в названии либо первый лист."""
    # Получаем список имён листов
    sheet_names = workbook.sheet_names
    for name in sheet_names:
        if "сплит" in name.lower():
            return workbook.get_sheet_by_name(name)
    if sheet_names:
        return workbook.get_sheet_by_name(sheet_names[0])
    return None


# --------------------------------------------------------------------------
# Выгрузка справочника цен в Excel
# --------------------------------------------------------------------------
def export_split_prices_to_excel(records, output_path: str,
                                 title: str = "Справочник цен (сплит-форма)") -> str:
    """Выгружает позиции справочника цен в Excel-файл.

    Args:
        records: список словарей (или строк DataFrame) с ключами code, name,
            unit, price_no_vat, price_vat, kind, group_code, group_name.
        output_path (str): путь к создаваемому .xlsx.
        title (str): название листа.

    Returns:
        str: путь к созданному файлу.
    """
    import xlsxwriter

    # DataFrame приводит к списку словарей: иначе обход даёт имена колонок
    if hasattr(records, 'to_dict'):
        records = records.to_dict('records')

    def get(rec, key):
        try:
            value = rec[key]
        except (KeyError, IndexError, TypeError):
            return ""
        if isinstance(value, float) and math.isnan(value):
            return ""
        return value

    wb = xlsxwriter.Workbook(output_path, {"constant_memory": True})
    try:
        ws = wb.add_worksheet("Цены")

        f_header = wb.add_format({
            "bold": True, "border": 1, "bg_color": "#D9E1F2", "align": "center",
            "valign": "vcenter", "text_wrap": True, "font_name": "Arial", "font_size": 11,
        })
        f_text = wb.add_format({"border": 1, "font_name": "Arial", "font_size": 10})
        f_num = wb.add_format({
            "border": 1, "num_format": "#,##0.00", "align": "right",
            "font_name": "Arial", "font_size": 10,
        })

        headers = [
            "Код", "Наименование", "Ед. изм.", "Цена без НДС 22%, руб.",
            "Цена с НДС 22%, руб.", "Тип", "Группа (номер)", "Группа (наименование)",
        ]
        for c, h in enumerate(headers):
            ws.write(0, c, h, f_header)

        ws.set_column(0, 0, 18)
        ws.set_column(1, 1, 70)
        ws.set_column(2, 2, 10)
        ws.set_column(3, 4, 18)
        ws.set_column(5, 5, 12)
        ws.set_column(6, 6, 12)
        ws.set_column(7, 7, 45)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, 0, len(headers) - 1)

        for r, rec in enumerate(records, start=1):
            ws.write_string(r, 0, str(get(rec, "code")), f_text)
            ws.write_string(r, 1, str(get(rec, "name")), f_text)
            ws.write_string(r, 2, str(get(rec, "unit")), f_text)
            ws.write_number(r, 3, float(get(rec, "price_no_vat") or 0), f_num)
            ws.write_number(r, 4, float(get(rec, "price_vat") or 0), f_num)
            ws.write_string(r, 5, str(get(rec, "kind")), f_text)
            ws.write_string(r, 6, str(get(rec, "group_code")), f_text)
            ws.write_string(r, 7, str(get(rec, "group_name")), f_text)
    finally:
        wb.close()

    return output_path