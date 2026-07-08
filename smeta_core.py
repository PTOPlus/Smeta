# -*- coding: utf-8 -*-
"""
smeta_core.py
Бизнес-логика приложения "Сметчик PRO", не зависящая от tkinter.
"""
import os
import math
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell as RC

COLS = ['Работа', 'Ед_изм_раб', 'Материал', 'Ед_изм',
        'Расход_1', 'Цена_мат_1', 'Цена_раб_1',
        'Расход_2', 'Цена_мат_2', 'Цена_раб_2']
LEGACY_COLS = ['Работа', 'Ед_изм_раб', 'Материал', 'Ед_изм', 'Расход', 'Цена_мат', 'Цена_раб']
CALC_HEADERS = ("№", "Наименование", "Ед. изм.",
                "Норма В1", "Объём В1", "Цена В1", "Стоимость В1",
                "Норма В2", "Объём В2", "Цена В2", "Стоимость В2")
EMPTY_TOKENS = ("", "-", "0", "nan", "none", "NaN")

# --------------------------------------------------------------------------
# Базовые помощники
# --------------------------------------------------------------------------
def to_float(val, default=0.0):
    """Надёжно парсит число, поддерживает запятую как десятичный разделитель."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        if isinstance(val, float) and math.isnan(val):
            return default
        return float(val)
    # ✅ ИСПРАВЛЕНО: убран лишний пробел перед \xa0
    s = str(val).strip().replace('\xa0', '').replace(' ', '').replace(',', '.')
    if s == "" or s.lower() in ("nan", "none", "-"):
        return default
    try:
        return float(s)
    except ValueError:
        return default

def is_section(name) -> bool:
    return str(name).strip().startswith("РАЗДЕЛ:")

def is_work(name) -> bool:
    return str(name).strip().startswith("Работа:")

def is_total(name) -> bool:
    return str(name).strip().startswith("ИТОГО ПО УЗЛУ")

def is_material(name) -> bool:
    return str(name).strip().startswith(">")

def clean_name(name) -> str:
    """Убирает служебные префиксы ('Работа:', '>', пробелы) из имени."""
    s = str(name).strip()
    if s.startswith("Работа: "):
        s = s[len("Работа: "):]
    elif s.startswith(">"):
        s = s[1:]
    return s.strip()

# --------------------------------------------------------------------------
# Работа с базой (DataFrame)
# --------------------------------------------------------------------------
def migrate_legacy_df(df):
    """Если в базе старый формат (один вариант цены) — переносит значения в '_1' колонки и копирует их же в '_2'."""
    has_legacy = all(c in df.columns for c in LEGACY_COLS)
    has_new = all(c in df.columns for c in COLS)

    if has_new:
        for c in COLS:
            if c not in df.columns:
                df[c] = 0.0 if c.startswith(('Расход', 'Цена')) else "-"
        return df[COLS].copy()

    if has_legacy:
        out = df.copy()
        out['Расход_1'] = out['Расход']
        out['Цена_мат_1'] = out['Цена_мат']
        out['Цена_раб_1'] = out['Цена_раб']
        out['Расход_2'] = out['Расход']
        out['Цена_мат_2'] = out['Цена_мат']
        out['Цена_раб_2'] = out['Цена_раб']
        return out[COLS].copy()

    for c in COLS:
        if c not in df.columns:
            df[c] = 0.0 if c.startswith(('Расход', 'Цена')) else "-"
    return df[COLS].copy()

def build_work_block(work_name, vol, db, next_num, db_manager=None):
    """Строит строки сметы (работа + материалы + строка ИТОГО ПО УЗЛУ).
    Если db_manager предоставлен — использует нормализованную БД."""
    if db_manager is not None:
        work_data = db_manager.get_work_with_materials(work_name)
        if work_data is None:
            return None
        
        work = work_data['work']
        materials = work_data['materials']
        
        unit_w = str(work['unit'])
        price_w1 = to_float(work['price_1'])
        price_w2 = to_float(work['price_2'])
        vol = to_float(vol)
        
        work_cost1 = round(vol * price_w1, 2)
        work_cost2 = round(vol * price_w2, 2)
        
        rows = [(next_num, f"Работа: {work_name}", unit_w, "-", vol, price_w1, work_cost1,
                 "-", vol, price_w2, work_cost2)]
        
        mat_total1 = 0.0
        mat_total2 = 0.0
        for mat in materials:
            mat_name = str(mat['name']).strip()
            if mat_name.lower() in EMPTY_TOKENS:
                continue
            
            rashod1 = to_float(mat['consumption_1'])
            price_m1 = to_float(mat['price_1'])
            rashod2 = to_float(mat['consumption_2'])
            price_m2 = to_float(mat['price_2'])
            
            qty1 = round(rashod1 * vol, 3)
            qty2 = round(rashod2 * vol, 3)
            cost1 = round(qty1 * price_m1, 2)
            cost2 = round(qty2 * price_m2, 2)
            
            mat_total1 += cost1
            mat_total2 += cost2
            
            rows.append(("", f"    > {mat_name}", str(mat['unit']), 
                        rashod1, qty1, price_m1, cost1,
                        rashod2, qty2, price_m2, cost2))
        
        combined1 = round(work_cost1 + mat_total1, 2)
        combined2 = round(work_cost2 + mat_total2, 2)
        rows.append(("", f"ИТОГО ПО УЗЛУ: {work_name}", "", "", "", "Сумма:", combined1,
                     "", "", "Сумма:", combined2))
        return rows
    else:
        # Старый метод с DataFrame
        items = db[db['Работа'].astype(str).str.strip() == str(work_name).strip()]
        if items.empty:
            return None
        first = items.iloc[0]
        unit_w = str(first['Ед_изм_раб'])
        price_w1 = to_float(first['Цена_раб_1'])
        price_w2 = to_float(first['Цена_раб_2'])
        vol = to_float(vol)
        
        work_cost1 = round(vol * price_w1, 2)
        work_cost2 = round(vol * price_w2, 2)
        
        rows = [(next_num, f"Работа: {work_name}", unit_w, "-", vol, price_w1, work_cost1,
                 "-", vol, price_w2, work_cost2)]
        
        mat_total1 = 0.0
        mat_total2 = 0.0
        for _, r in items.iterrows():
            mat_name = str(r['Материал']).strip()
            if mat_name.lower() in EMPTY_TOKENS:
                continue
            rashod1 = to_float(r['Расход_1'])
            price_m1 = to_float(r['Цена_мат_1'])
            rashod2 = to_float(r['Расход_2'])
            price_m2 = to_float(r['Цена_мат_2'])
            qty1 = round(rashod1 * vol, 3)
            qty2 = round(rashod2 * vol, 3)
            cost1 = round(qty1 * price_m1, 2)
            cost2 = round(qty2 * price_m2, 2)
            mat_total1 += cost1
            mat_total2 += cost2
            rows.append(("", f"    > {mat_name}", str(r['Ед_изм']), rashod1, qty1, price_m1, cost1,
                         rashod2, qty2, price_m2, cost2))
        
        combined1 = round(work_cost1 + mat_total1, 2)
        combined2 = round(work_cost2 + mat_total2, 2)
        rows.append(("", f"ИТОГО ПО УЗЛУ: {work_name}", "", "", "", "Сумма:", combined1,
                     "", "", "Сумма:", combined2))
        return rows

def rebuild_smeta(rows):
    """Полная пересборка сметы — единая точка пересчёта."""
    cleaned = [r for r in rows if not is_total(str(r[1]))]
    out = []
    next_num = 1
    work_vol = 1.0
    node_work_name = None
    node_work_cost1 = node_work_cost2 = 0.0
    node_mat_cost1 = node_mat_cost2 = 0.0

    def flush_node():
        nonlocal node_work_name, node_work_cost1, node_work_cost2, node_mat_cost1, node_mat_cost2
        if node_work_name is not None:
            out.append(("", f"ИТОГО ПО УЗЛУ: {node_work_name}", "", "", "", "Сумма:",
                        round(node_work_cost1 + node_mat_cost1, 2),
                        "", "", "Сумма:", round(node_work_cost2 + node_mat_cost2, 2)))
        node_work_name = None
        node_work_cost1 = node_work_cost2 = node_mat_cost1 = node_mat_cost2 = 0.0

    for raw in cleaned:
        vals = list(raw)
        name = str(vals[1]).strip()

        if is_section(name):
            flush_node()
            out.append(tuple(vals))
            continue

        if is_work(name):
            flush_node()
            user_vol = to_float(vals[4], 0.0)
            work_vol = user_vol if user_vol > 0 else 1.0
            
            price1 = to_float(vals[5])
            price2 = to_float(vals[9])
            vals[0] = next_num
            next_num += 1
            vals[4] = work_vol
            vals[8] = work_vol
            vals[6] = round(work_vol * price1, 2)
            vals[10] = round(work_vol * price2, 2)
            node_work_name = name.replace("Работа: ", "").strip()
            node_work_cost1, node_work_cost2 = vals[6], vals[10]
            out.append(tuple(vals))
            continue

        if is_material(name):
            # Объём материала — всегда производная величина (Норма × Объём работы).
            # Колонки "Объём" (индексы 4 и 8) НЕ входят в список редактируемых
            # для строки материала (см. _editable_cols_for в app.py — там только
            # {1, 3, 5, 7, 9}), то есть пользователь никогда не вводит их вручную.
            # Старый код вместо этого проверял "если текущее значение > 0 — не
            # трогать", из-за чего при изменении Объёма РАБОТЫ материалы не
            # пересчитывались и оставались с прежним (уже устаревшим) объёмом.
            # Проверено тестом: Работа 10->50, Материал с нормой 2 должен дать
            # 100, а не оставаться на 20.
            norm1 = to_float(vals[3])
            norm2 = to_float(vals[7])
            vals[4] = round(norm1 * work_vol, 3)
            vals[8] = round(norm2 * work_vol, 3)

            price1 = to_float(vals[5])
            price2 = to_float(vals[9])
            vals[0] = ""
            vals[6] = round(vals[4] * price1, 2)
            vals[10] = round(vals[8] * price2, 2)
            node_mat_cost1 += vals[6]
            node_mat_cost2 += vals[10]
            out.append(tuple(vals))
            continue

        out.append(tuple(vals))

    flush_node()
    return out

def compute_grand_totals(rows):
    """Возвращает (итого_в1, итого_в2) — сумму всех строк 'ИТОГО ПО УЗЛУ'."""
    t1 = t2 = 0.0
    for vals in rows:
        name = str(vals[1]).strip()
        if is_total(name):
            t1 += to_float(vals[6])
            t2 += to_float(vals[10])
    return round(t1, 2), round(t2, 2)

# --------------------------------------------------------------------------
# Импорт ранее выгруженной сметы
# --------------------------------------------------------------------------
def parse_exported_sheet(sheet_values):
    """Разбирает ранее экспортированный лист 'Смета'."""
    import math
    
    title = ""
    overhead = (0.0, 0.0)
    lifting = (0.0, 0.0)
    lifting_trash = (0.0, 0.0)
    sequence = []

    if sheet_values:
        for row in sheet_values[:3]:
            for cell in row:
                if cell and isinstance(cell, str) and len(str(cell).strip()) > 20:
                    if "Смета" in cell or "смета" in cell or "Коммерческое" in cell:
                        title = str(cell).strip()
                        break

    for row_idx, row in enumerate(sheet_values): 
        if not row:
            continue

        col0 = row[0] if len(row) > 0 else None
        col1 = row[1] if len(row) > 1 else None
        
        def safe_str(val):
            if val is None:
                return ""
            if isinstance(val, float) and math.isnan(val):
                return ""
            return str(val).strip()
        
        str0 = safe_str(col0)
        str1 = safe_str(col1)
        
        if "№ п/п" in str0 or "№" in str0:
            continue
        if "Наименование" in str1 and "Ед. изм" in str(row[2] if len(row) > 2 else ""):
            continue
        if str0.isdigit() and str1.isdigit() and int(str0) < 10 and int(str1) < 10:
            continue
            
        if str1.startswith("РАЗДЕЛ:") or (str0 and not str0.replace('.', '').isdigit() and not str1):
            if str0 and len(str0) > 5 and not str0.replace('.', '').isdigit():
                sequence.append(('section', str0.replace("РАЗДЕЛ:", "").strip()))
                continue
        
        if not str1 and str0 and not str0.replace('.', '').isdigit() and len(str0) > 3:
            if "Гидроизоляция" in str0 or "Устройство" in str0 or "Монтаж" in str0:
                sequence.append(('section', str0))
                continue
        
        if not str1:
            continue
            
        name = str1
        if not name:
            continue

        if name.startswith("в т.ч.") or name.startswith("- материалы") or name.startswith("- работы"):
            continue
        if name.startswith("ИТОГО") or name.startswith("Итого"):
            continue
        if name == "Смета" or "на выполнение" in name:
            continue

        if "Накладные" in name or "транспортные" in name:
            try:
                val1 = float(row[6]) if len(row) > 6 and not (isinstance(row[6], float) and math.isnan(row[6])) else 0.0
                val2 = float(row[10]) if len(row) > 10 and not (isinstance(row[10], float) and math.isnan(row[10])) else 0.0
                overhead = (val1, val2)
            except (ValueError, TypeError, IndexError):
                pass
            continue
        
        if "Подъёмные" in name or "Подъемные" in name or "грузопод" in name.lower():
            try:
                val1 = float(row[6]) if len(row) > 6 and not (isinstance(row[6], float) and math.isnan(row[6])) else 0.0
                val2 = float(row[10]) if len(row) > 10 and not (isinstance(row[10], float) and math.isnan(row[10])) else 0.0
                lifting = (val1, val2)
            except (ValueError, TypeError, IndexError):
                pass
            continue
            
        if "вывоз" in name.lower() and "мусор" in name.lower():
            try:
                val1 = float(row[6]) if len(row) > 6 and not (isinstance(row[6], float) and math.isnan(row[6])) else 0.0
                val2 = float(row[10]) if len(row) > 10 and not (isinstance(row[10], float) and math.isnan(row[10])) else 0.0
                lifting_trash = (val1, val2)
            except (ValueError, TypeError, IndexError):
                pass
            continue

        is_numbered = False
        if str0 and str0.replace('.', '').isdigit():
            is_numbered = True

        if is_numbered:
            try:
                vol = float(row[4]) if len(row) > 4 and not (isinstance(row[4], float) and math.isnan(row[4])) else 0.0
            except (ValueError, TypeError, IndexError):
                vol = 0.0
            sequence.append(('work', name, vol))

    return {
        'title': title,
        'overhead': overhead,
        'lifting': lifting,
        'lifting_trash': lifting_trash,
        'sequence': sequence,
    }

# --------------------------------------------------------------------------
# Экспорт в Excel
# --------------------------------------------------------------------------
def export_smeta_to_excel(rows, output_path, title="", meta_rows=None,
                          overhead1=0.0, overhead2=0.0, lift1=0.0, lift2=0.0,
                          trash1=0.0, trash2=0.0):
    import xlsxwriter
    wb = xlsxwriter.Workbook(output_path)
    ws = wb.add_worksheet("Смета")

    f_title = wb.add_format({'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'font_name': 'Arial'})
    f_variant = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#BDD7EE', 'align': 'center', 'font_name': 'Arial', 'font_size': 11})
    f_header = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'font_name': 'Arial', 'font_size': 11})
    f_section = wb.add_format({'bold': True, 'bg_color': '#E1BEE7', 'border': 1, 'align': 'left', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 11})
    f_item_txt = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC', 'text_wrap': True, 'font_name': 'Arial', 'font_size': 10})
    f_item_ctr = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'font_name': 'Arial', 'font_size': 10})
    f_item_num = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC', 'num_format': '#,##0.00', 'align': 'right', 'font_name': 'Arial', 'font_size': 10})
    f_break_txt = wb.add_format({'italic': True, 'border': 1, 'align': 'right', 'font_name': 'Arial', 'font_size': 10})
    f_break_num = wb.add_format({'italic': True, 'border': 1, 'num_format': '#,##0.00', 'align': 'right', 'font_name': 'Arial', 'font_size': 10})
    f_mat_txt = wb.add_format({'border': 1, 'indent': 1, 'text_wrap': True, 'font_name': 'Arial', 'font_size': 10})
    f_mat_num = wb.add_format({'border': 1, 'num_format': '#,##0.00', 'align': 'right', 'font_name': 'Arial', 'font_size': 10})
    f_blank = wb.add_format({'border': 1, 'font_name': 'Arial', 'font_size': 10})
    f_sub_lbl = wb.add_format({'bold': True, 'border': 1, 'align': 'left', 'font_name': 'Arial', 'font_size': 11})
    f_sub_num = wb.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'align': 'right', 'font_name': 'Arial', 'font_size': 11})
    f_total_lbl = wb.add_format({'bold': True, 'border': 1, 'align': 'left', 'font_size': 12, 'font_name': 'Arial'})
    f_total_num = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#FFCC00', 'num_format': '#,##0.00', 'align': 'right', 'font_name': 'Arial', 'font_size': 12})
    f_ratio = wb.add_format({'border': 1, 'num_format': '0.000', 'align': 'right', 'font_name': 'Arial', 'font_size': 10})

    ws.merge_range(0, 0, 0, 10, title or "Смета", f_title)
    ws.merge_range(1, 3, 1, 6, "ЦЕНА — ВАРИАНТ 1", f_variant)
    ws.merge_range(1, 7, 1, 10, "ЦЕНА — ВАРИАНТ 2", f_variant)

    headers = ["№ п/п", "Наименование работ и затрат", "Ед. изм.",
               "Норма расхода", "Объём", "Цена за ед., руб.", "Сметная стоимость, руб.",
               "Норма расхода", "Объём", "Цена за ед., руб.", "Сметная стоимость, руб."]
    for c, h in enumerate(headers):
        ws.write(2, c, h, f_header)

    ws.set_column(0, 0, 6)
    ws.set_column(1, 1, 52)
    ws.set_column(2, 2, 9)
    ws.set_column(3, 3, 11)
    ws.set_column(4, 4, 11)
    ws.set_column(5, 5, 13)
    ws.set_column(6, 6, 16)
    ws.set_column(7, 7, 11)
    ws.set_column(8, 8, 11)
    ws.set_column(9, 9, 13)
    ws.set_column(10, 10, 16)
    ws.freeze_panes(3, 2)
 
    excel_row = 3
    i = 0
    n = len(rows)

    workonly_total1_refs, workonly_total2_refs = [], []
    matonly_total1_refs, matonly_total2_refs = [], []
    seen_work1, seen_work2 = {}, {}
    seen_mat1, seen_mat2 = {}, {}

    def price_cell(seen_dict, key, row_idx, col_idx, value, fmt):
        if key not in seen_dict:
            seen_dict[key] = (row_idx, col_idx)
            ws.write(row_idx, col_idx, value, fmt)
        else:
            ref_row, ref_col = seen_dict[key]
            ws.write_formula(row_idx, col_idx, f"={RC(ref_row, ref_col)}", fmt, value)

    while i < n:
        vals = rows[i]
        name = str(vals[1]).strip()

        if name.startswith("РАЗДЕЛ: "):
            section_title = name.replace("РАЗДЕЛ: ", "").strip()
            ws.merge_range(excel_row, 0, excel_row, 10, section_title, f_section)
            excel_row += 1
            i += 1
            continue

        if name.startswith("Работа: "):
            work_name = name.replace("Работа: ", "").strip()
            unit_w = str(vals[2])
            try:
                work_num = int(vals[0])
            except (TypeError, ValueError):
                work_num = ""
            vol = to_float(vals[4])
            price_w1 = to_float(vals[5])
            price_w2 = to_float(vals[9])

            j = i + 1
            mats = []
            while j < n and str(rows[j][1]).strip().startswith(">"):
                mats.append(rows[j])
                j += 1

            row_top = excel_row
            row_workonly = excel_row + 1
            row_matonly = excel_row + 2
            row_mat_start = excel_row + 3
            row_mat_end = row_mat_start + len(mats) - 1

            ws.write(row_top, 0, work_num, f_item_ctr)
            ws.write(row_top, 1, work_name, f_item_txt)
            ws.write(row_top, 2, unit_w, f_item_ctr)
            ws.write_blank(row_top, 3, None, f_item_txt)
            ws.write(row_top, 4, vol, f_item_num)
            ws.write(row_top, 5, price_w1, f_item_num)
            ws.write_formula(row_top, 6, f"={RC(row_workonly, 6)}+{RC(row_matonly, 6)}", f_item_num)
            ws.write_blank(row_top, 7, None, f_item_txt)
            ws.write(row_top, 8, vol, f_item_num)
            ws.write(row_top, 9, price_w2, f_item_num)
            ws.write_formula(row_top, 10, f"={RC(row_workonly, 10)}+{RC(row_matonly, 10)}", f_item_num)

            ws.write_blank(row_workonly, 0, None, f_blank)
            ws.write(row_workonly, 1, "в т.ч.: - работы:", f_break_txt)
            ws.write_blank(row_workonly, 2, None, f_blank)
            ws.write_blank(row_workonly, 3, None, f_blank)
            ws.write_blank(row_workonly, 4, None, f_blank)  
            price_cell(seen_work1, work_name, row_workonly, 5, price_w1, f_break_num)
            ws.write_formula(row_workonly, 6, f"={RC(row_top, 4)}*{RC(row_workonly, 5)}", f_break_num)
            ws.write_blank(row_workonly, 7, None, f_blank)
            ws.write_blank(row_workonly, 8, None, f_blank)
            price_cell(seen_work2, work_name, row_workonly, 9, price_w2, f_break_num)
            ws.write_formula(row_workonly, 10, f"={RC(row_top, 8)}*{RC(row_workonly, 9)}", f_break_num)
            workonly_total1_refs.append(RC(row_workonly, 6))
            workonly_total2_refs.append(RC(row_workonly, 10))

            ws.write_blank(row_matonly, 0, None, f_blank)
            ws.write(row_matonly, 1, "- материалы:", f_break_txt)
            for c in (2, 3, 4, 5, 7, 8, 9):
                ws.write_blank(row_matonly, c, None, f_blank)
            if mats:
                ws.write_formula(row_matonly, 6, f"=SUM({RC(row_mat_start, 6)}:{RC(row_mat_end, 6)})", f_break_num)
                ws.write_formula(row_matonly, 10, f"=SUM({RC(row_mat_start, 10)}:{RC(row_mat_end, 10)})", f_break_num)
            else:
                ws.write(row_matonly, 6, 0, f_break_num)
                ws.write(row_matonly, 10, 0, f_break_num)
            matonly_total1_refs.append(RC(row_matonly, 6))
            matonly_total2_refs.append(RC(row_matonly, 10))

            for k, mvals in enumerate(mats):
                r = row_mat_start + k
                mat_name = clean_name(str(mvals[1]))
                unit_m = str(mvals[2])
                norm1 = to_float(mvals[3])
                norm2 = to_float(mvals[7])

                ws.write_blank(r, 0, None, f_blank)
                ws.write(r, 1, mat_name, f_mat_txt)
                ws.write(r, 2, unit_m, f_mat_txt)
                ws.write(r, 3, norm1, f_mat_num)
                ws.write_formula(r, 4, f"={RC(r, 3)}*{RC(row_top, 4)}", f_mat_num)
                price_cell(seen_mat1, mat_name, r, 5, to_float(mvals[5]), f_mat_num)
                ws.write_formula(r, 6, f"={RC(r, 4)}*{RC(r, 5)}", f_mat_num)
                ws.write(r, 7, norm2, f_mat_num)
                ws.write_formula(r, 8, f"={RC(r, 7)}*{RC(row_top, 8)}", f_mat_num)
                price_cell(seen_mat2, mat_name, r, 9, to_float(mvals[9]), f_mat_num)
                ws.write_formula(r, 10, f"={RC(r, 8)}*{RC(r, 9)}", f_mat_num)

            excel_row = row_mat_end + 1
            i = j
        else:
            i += 1

    if meta_rows:
        meta_df = pd.DataFrame(meta_rows, columns=COLS).drop_duplicates().reset_index(drop=True)
        if not meta_df.empty:
            ws_meta = wb.add_worksheet("Meta")
            for c, h in enumerate(COLS):
                ws_meta.write(0, c, h)
            for r_idx, row in enumerate(meta_df.itertuples(index=False), start=1):
                for c_idx, val in enumerate(row):
                    ws_meta.write(r_idx, c_idx, val)

    row_gap = excel_row + 1
    row_total = row_gap + 1
    row_works = row_total + 1
    row_mats = row_total + 2
    row_overhead = row_total + 3
    row_lifting = row_total + 4
    row_trash = row_total + 5

    works_f1 = "+".join(workonly_total1_refs) if workonly_total1_refs else "0"
    works_f2 = "+".join(workonly_total2_refs) if workonly_total2_refs else "0"
    mats_f1 = "+".join(matonly_total1_refs) if matonly_total1_refs else "0"
    mats_f2 = "+".join(matonly_total2_refs) if matonly_total2_refs else "0"

    ws.write(row_works, 1, "в т.ч.: - работы:", f_sub_lbl)
    ws.write_formula(row_works, 6, f"={works_f1}", f_sub_num)
    ws.write_formula(row_works, 10, f"={works_f2}", f_sub_num)
    ws.write_formula(row_works, 9, f"=IF({RC(row_works,6)}=0,0,{RC(row_works,10)}/{RC(row_works,6)})", f_ratio)

    ws.write(row_mats, 1, "- материалы:", f_sub_lbl)
    ws.write_formula(row_mats, 6, f"={mats_f1}", f_sub_num)
    ws.write_formula(row_mats, 10, f"={mats_f2}", f_sub_num)
    ws.write_formula(row_mats, 9, f"=IF({RC(row_mats,6)}=0,0,{RC(row_mats,10)}/{RC(row_mats,6)})", f_ratio)

    ws.write(row_overhead, 1, "Накладные и транспортные расходы", f_sub_lbl)
    ws.write(row_overhead, 6, overhead1, f_sub_num)
    ws.write(row_overhead, 10, overhead2, f_sub_num)
    ws.write_formula(row_overhead, 9, f"=IF({RC(row_overhead,6)}=0,0,{RC(row_overhead,10)}/{RC(row_overhead,6)})", f_ratio)

    ws.write(row_lifting, 1, "Подъёмные механизмы", f_sub_lbl)
    ws.write(row_lifting, 6, lift1, f_sub_num)
    ws.write(row_lifting, 10, lift2, f_sub_num)
    ws.write_formula(row_lifting, 9, f"=IF({RC(row_lifting,6)}=0,0,{RC(row_lifting,10)}/{RC(row_lifting,6)})", f_ratio)

    ws.write(row_trash, 1, "Вывоз мусора", f_sub_lbl)
    ws.write(row_trash, 6, trash1, f_sub_num)
    ws.write(row_trash, 10, trash2, f_sub_num)
    ws.write_formula(row_trash, 9, f"=IF({RC(row_trash,6)}=0,0,{RC(row_trash,10)}/{RC(row_trash,6)})", f_ratio)

    ws.write(row_total, 1, "ИТОГО:", f_total_lbl)
    total_f1 = f"={RC(row_works,6)}+{RC(row_mats,6)}+{RC(row_overhead,6)}+{RC(row_lifting,6)}+{RC(row_trash,6)}"
    total_f2 = f"={RC(row_works,10)}+{RC(row_mats,10)}+{RC(row_overhead,10)}+{RC(row_lifting,10)}+{RC(row_trash,10)}"
    ws.write_formula(row_total, 6, total_f1, f_total_num)
    ws.write_formula(row_total, 10, total_f2, f_total_num)
    ws.write_formula(row_total, 8, f"={RC(row_total,6)}-{RC(row_total,10)}", f_total_num)
    ws.write_formula(row_total, 9, f"=IF({RC(row_total,6)}=0,0,{RC(row_total,10)}/{RC(row_total,6)})", f_ratio)

    wb.close()
    return output_path