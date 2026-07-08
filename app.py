# -*- coding: utf-8 -*-
"""
Сметчик PRO 5.1
Изменения: поддержка двух вариантов цены, сравнение экономии,
учёт доп. расходов, корректный импорт/экспорт, фикс ручных правок,
нормализованная база данных через Parquet.
"""
import os
import re
import json
import math
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import smeta_core as sc

# ✅ ИМПОРТ МЕНЕДЖЕРА БАЗЫ ДАННЫХ
try:
    from db_manager import DatabaseManager
    HAS_DB_MANAGER = True
except ImportError:
    HAS_DB_MANAGER = False

SETTINGS_FILE = 'settings.json'

# --------------------------------------------------------------------------
# Настройки
# --------------------------------------------------------------------------
def load_settings():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_settings = {
        'db_folder': script_dir,
        'export_folder': script_dir,
        'active_db_filename': 'my_works_base.xlsx',
        'auto_price_mat_ratio': 1.35,
        'auto_price_work_ratio': 2.2,
        'db_col_widths': {},
        'calc_col_widths': {},
    }
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            for key in default_settings:
                settings.setdefault(key, default_settings[key])
            return settings
        except Exception as e:
            print(f"Ошибка загрузки настроек: {e}")
    return default_settings

def save_settings(settings):
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Ошибка сохранения настроек: {e}")

# --------------------------------------------------------------------------
# Буфер обмена / контекстное меню
# --------------------------------------------------------------------------
def add_clipboard_support(widget):
    def copy(event):
        try:
            text = event.widget.selection_get()
            event.widget.clipboard_clear()
            event.widget.clipboard_append(text)
        except tk.TclError:
            pass
        return "break"
    
    def paste(event):
        try:
            event.widget.insert(tk.INSERT, event.widget.clipboard_get())
        except tk.TclError:
            pass
        return "break"
    
    def cut(event=None):
        try:
            text = event.widget.selection_get()
            event.widget.clipboard_clear()
            event.widget.clipboard_append(text)
            event.widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            pass
        return "break"
    
    def handle_ctrl_key(event):
        # Определяем модификатор: Control (Windows/Linux) или Command (macOS)
        is_ctrl = (event.state & 0x4) != 0   # Control
        is_cmd = (event.state & 0x10) != 0   # Command (macOS)
        if not (is_ctrl or is_cmd):
            return
        
        # Keycode не зависит от раскладки!
        if event.keycode == 67:   # C
            return copy(event)
        elif event.keycode == 86: # V
            return paste(event)
        elif event.keycode == 88: # X
            return cut(event)
        # Z (undo) обрабатывается отдельно в tree_smeta
        
    # Привязываем ко всем нажатиям с Control/Command
    widget.bind("<Control-Key>", handle_ctrl_key, add="+")
    widget.bind("<Command-Key>", handle_ctrl_key, add="+")
    
    # Оставляем старые привязки как fallback (на случай нестандартных клавиатур)
    widget.bind("<Control-c>", copy, add="+")
    widget.bind("<Control-v>", paste, add="+")
    widget.bind("<Command-c>", copy, add="+")
    widget.bind("<Command-v>", paste, add="+")

def add_context_menu(widget):
    menu = tk.Menu(widget, tearoff=0)
    def cut():
        try:
            text = widget.selection_get()
            widget.clipboard_clear()
            widget.clipboard_append(text)
            widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            pass
    def copy():
        try:
            text = widget.selection_get()
            widget.clipboard_clear()
            widget.clipboard_append(text)
        except tk.TclError:
            pass
    def paste():
        try:
            widget.insert(tk.INSERT, widget.clipboard_get())
        except tk.TclError:
            pass
    menu.add_command(label="Вырезать", command=cut)
    menu.add_command(label="Копировать", command=copy)
    menu.add_command(label="Вставить", command=paste)
    def show_menu(event):
        try:
            has_selection = widget.tag_ranges(tk.SEL)
        except Exception:
            has_selection = False
        menu.entryconfigure(0, state="normal" if has_selection else "disabled")
        menu.entryconfigure(1, state="normal" if has_selection else "disabled")
        try:
            widget.clipboard_get()
            menu.entryconfigure(2, state="normal")
        except tk.TclError:
            menu.entryconfigure(2, state="disabled")
        menu.tk_popup(event.x_root, event.y_root)
    widget.bind("<Button-3>", show_menu)

# --------------------------------------------------------------------------
# Главный класс приложения
# --------------------------------------------------------------------------
class SmetaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Сметчик PRO 5.1")
        self.root.geometry("1650x900")
        
        settings = load_settings()
        self.settings = settings
        self.db_folder = settings['db_folder']
        self.export_folder = settings['export_folder']
        self.active_db_filename = settings.get('active_db_filename', 'my_works_base.xlsx')
        self.db_file = os.path.join(self.db_folder, self.active_db_filename)

        self.edit_entry = None
        self.edit_item = None
        self.edit_col_idx = None
        self.edit_orig = None
        self.undo_stack = []
        self.ctx_menu_item = None
        self.ctx_menu_col = None

        # ✅ ИНИЦИАЛИЗАЦИЯ МЕНЕДЖЕРА БД
        if HAS_DB_MANAGER:
            db_name = self.active_db_filename.rsplit('.', 1)[0]
            self.db_manager = DatabaseManager(self.db_folder, db_name)
        else:
            self.db_manager = None

        self.db = self._load_db()
        self.sort_orders = {col: False for col in sc.COLS}
        self.display_cols = ['Работа', 'Ед_изм_раб', 'Цена_раб_1', 'Цена_раб_2',
                             'Материал', 'Ед_изм', 'Расход_1', 'Цена_мат_1', 'Расход_2', 'Цена_мат_2']

        menubar = tk.Menu(root)
        root.config(menu=menubar)
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Настройки", menu=settings_menu)
        settings_menu.add_command(label="Настройки", command=self.open_settings)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Справка", menu=help_menu)
        help_menu.add_command(label="О программе", command=self.show_about)

        self.notebook = ttk.Notebook(root)
        self.tab_calc = tk.Frame(self.notebook)
        self.tab_db = tk.Frame(self.notebook)
        self.notebook.add(self.tab_calc, text="Составление сметы")
        self.notebook.add(self.tab_db, text="Справочник")
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.setup_db_tab()
        self.setup_calc_tab()
        self.apply_column_widths()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.refresh_db_list()

    def show_about(self):
        messagebox.showinfo("О программе", "Сметчик PRO 5.1\n\nПоддержка двух вариантов цены, учёт доп. расходов, автоматический расчёт экономии, нормализованная база данных.")

    def open_settings(self):
        settings_win = tk.Toplevel(self.root)
        settings_win.title("Настройки")
        settings_win.geometry("600x380")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()
        db_folder_var = tk.StringVar(value=self.db_folder)
        export_folder_var = tk.StringVar(value=self.export_folder)
        mat_ratio_var = tk.StringVar(value=str(self.settings.get('auto_price_mat_ratio', 1.35)))
        work_ratio_var = tk.StringVar(value=str(self.settings.get('auto_price_work_ratio', 2.2)))
        frame_db = tk.LabelFrame(settings_win, text="Папка для базы данных", padx=10, pady=10)
        frame_db.pack(fill=tk.X, padx=10, pady=5)
        tk.Entry(frame_db, textvariable=db_folder_var, width=60).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_db, text="Обзор...", command=lambda: self.browse_folder(db_folder_var)).pack(side=tk.LEFT)
        frame_export = tk.LabelFrame(settings_win, text="Папка для сохранения смет", padx=10, pady=10)
        frame_export.pack(fill=tk.X, padx=10, pady=5)
        tk.Entry(frame_export, textvariable=export_folder_var, width=60).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_export, text="Обзор...", command=lambda: self.browse_folder(export_folder_var)).pack(side=tk.LEFT)
        frame_ratios = tk.LabelFrame(settings_win, text="Коэффициенты автоподстановки цены В1 от В2", padx=10, pady=10)
        frame_ratios.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame_ratios, text="Материалы (В1 = В2 ×)").grid(row=0, column=0, sticky="w", pady=2)
        mat_ratio_entry = tk.Entry(frame_ratios, textvariable=mat_ratio_var, width=10)
        add_clipboard_support(mat_ratio_entry)
        mat_ratio_entry.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        tk.Label(frame_ratios, text="(например, 1.35)").grid(row=0, column=2, sticky="w", padx=5)
        tk.Label(frame_ratios, text="Работы (В1 = В2 ×)").grid(row=1, column=0, sticky="w", pady=2)
        work_ratio_entry = tk.Entry(frame_ratios, textvariable=work_ratio_var, width=10)
        add_clipboard_support(work_ratio_entry)
        work_ratio_entry.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        tk.Label(frame_ratios, text="(например, 2.2)").grid(row=1, column=2, sticky="w", padx=5)
        btn_frame = tk.Frame(settings_win)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="Сохранить", bg="#4CAF50", fg="white", width=15,
                command=lambda: self.save_settings_and_close(
                    settings_win, db_folder_var.get(), export_folder_var.get(),
                    mat_ratio_var.get(), work_ratio_var.get()
                )).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Отмена", width=15, command=settings_win.destroy).pack(side=tk.LEFT, padx=10)

    def browse_folder(self, var):
        folder = filedialog.askdirectory(title="Выберите папку")
        if folder:
            var.set(folder)

    def save_settings_and_close(self, win, db_folder, export_folder, mat_ratio_str, work_ratio_str):
        os.makedirs(db_folder, exist_ok=True)
        os.makedirs(export_folder, exist_ok=True)
        try:
            mat_ratio = float(mat_ratio_str.replace(',', '.'))
            work_ratio = float(work_ratio_str.replace(',', '.'))
            if mat_ratio <= 0 or work_ratio <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Ошибка", "Коэффициенты должны быть положительными числами.")
            return
        self.db_folder = db_folder
        self.export_folder = export_folder
        self.db_file = os.path.join(self.db_folder, self.active_db_filename)
        self.settings['auto_price_mat_ratio'] = mat_ratio
        self.settings['auto_price_work_ratio'] = work_ratio
        save_settings({
            'db_folder': db_folder,
            'export_folder': export_folder,
            'active_db_filename': self.active_db_filename,
            'auto_price_mat_ratio': mat_ratio,
            'auto_price_work_ratio': work_ratio,
        })
        self.db = self._load_db()
        self.refresh_db_table()
        self.update_combobox()
        self.refresh_db_list()
        messagebox.showinfo("Готово", "Настройки сохранены.")
        win.destroy()

    def save_column_widths(self):
        """Сохраняет текущие ширины колонок в настройки."""
        db_widths = {}
        if hasattr(self, 'tree_db'):
            for col in self.display_cols:
                try:
                    db_widths[col] = self.tree_db.column(col)['width']
                except Exception:
                    pass
        
        calc_widths = {}
        if hasattr(self, 'tree_smeta'):
            for col in self.calc_cols:
                try:
                    calc_widths[col] = self.tree_smeta.column(col)['width']
                except Exception:
                    pass
        
        self.settings['db_col_widths'] = db_widths
        self.settings['calc_col_widths'] = calc_widths
        save_settings(self.settings)

    def apply_column_widths(self):
        """Применяет сохранённые ширины колонок."""
        db_widths = self.settings.get('db_col_widths', {})
        if hasattr(self, 'tree_db'):
            for col, width in db_widths.items():
                try:
                    self.tree_db.column(col, width=width)
                except Exception:
                    pass
        
        calc_widths = self.settings.get('calc_col_widths', {})
        if hasattr(self, 'tree_smeta'):
            for col, width in calc_widths.items():
                try:
                    self.tree_smeta.column(col, width=width)
                except Exception:
                    pass

    def on_closing(self):
        """Обработчик закрытия окна — сохраняет ширины колонок."""
        self.save_column_widths()
        self.root.destroy()

    def _load_db(self):
        """Загружает базу данных (с поддержкой Parquet и Excel)."""
        if self.db_manager is not None:
            return self.db_manager.get_legacy_dataframe()
        else:
            if not os.path.exists(self.db_file):
                return pd.DataFrame(columns=sc.COLS)
            try:
                raw = pd.read_excel(self.db_file)
            except Exception as e:
                messagebox.showerror("Ошибка БД", f"Не удалось загрузить базу:\n{e}")
                return pd.DataFrame(columns=sc.COLS)
            is_legacy_only = (all(c in raw.columns for c in sc.LEGACY_COLS)
                            and not all(c in raw.columns for c in sc.COLS))
            migrated = sc.migrate_legacy_df(raw)
            if is_legacy_only:
                try:
                    backup_path = self.db_file.rsplit('.', 1)[0] + "_backup_old_format.xlsx"
                    if not os.path.exists(backup_path):
                        raw.to_excel(backup_path, index=False)
                    migrated.to_excel(self.db_file, index=False)
                    messagebox.showinfo("База обновлена", 
                        f"Формат обновлён под два варианта цены.\n"
                        f"Резервная копия: {os.path.basename(backup_path)}")
                except Exception as e:
                    messagebox.showwarning("Внимание", 
                        f"База мигрирована в памяти, но не удалось сохранить файл:\n{e}")
            return migrated

    def setup_db_tab(self):
        db_sel_frame = tk.Frame(self.tab_db)
        db_sel_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(db_sel_frame, text="📂 База данных:").pack(side=tk.LEFT)
        self.db_combo = ttk.Combobox(db_sel_frame, width=40, state="readonly")
        self.db_combo.pack(side=tk.LEFT, padx=5)
        tk.Button(db_sel_frame, text="➕ Создать", command=self.create_new_db).pack(side=tk.LEFT, padx=2)
        tk.Button(db_sel_frame, text="🗑 Удалить", command=self.delete_current_db).pack(side=tk.LEFT, padx=2)
        tk.Button(db_sel_frame, text="🔄 Обновить", command=self.refresh_db_list).pack(side=tk.LEFT, padx=2)
        self.db_combo.bind("<<ComboboxSelected>>", self.on_db_selected)

        frame_input = tk.LabelFrame(self.tab_db, text="Редактор базы", padx=10, pady=10)
        frame_input.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame_input, text="Наименование работы:").grid(row=0, column=0, sticky="nw")
        self.work_text = tk.Text(frame_input, width=110, height=2, font=("Arial", 10))
        add_clipboard_support(self.work_text)
        add_context_menu(self.work_text)
        self.work_text.grid(row=0, column=0, columnspan=4, padx=5, pady=(20, 5), sticky="w")
        self.entries = {}
        tk.Label(frame_input, text="Ед. изм. работы:").grid(row=1, column=0, sticky="w", pady=2)
        e = tk.Entry(frame_input, width=18)
        add_clipboard_support(e)
        add_context_menu(e)
        e.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        self.entries['Ед_изм_раб'] = e
        tk.Label(frame_input, text="Материал:").grid(row=1, column=2, sticky="w", pady=2)
        e = tk.Entry(frame_input, width=40)
        add_clipboard_support(e)
        add_context_menu(e)
        e.grid(row=1, column=3, padx=5, pady=2, sticky="w")
        self.entries['Материал'] = e
        tk.Label(frame_input, text="Ед. изм. материала:").grid(row=2, column=0, sticky="w", pady=2)
        e = tk.Entry(frame_input, width=18)
        add_clipboard_support(e)
        add_context_menu(e)
        e.grid(row=2, column=1, padx=5, pady=2, sticky="w")
        self.entries['Ед_изм'] = e
        frame_v1 = tk.LabelFrame(frame_input, text="Вариант 1", padx=8, pady=6)
        frame_v1.grid(row=3, column=0, columnspan=2, padx=5, pady=10, sticky="we")
        frame_v2 = tk.LabelFrame(frame_input, text="Вариант 2", padx=8, pady=6)
        frame_v2.grid(row=3, column=2, columnspan=2, padx=5, pady=10, sticky="we")
        variant_fields = [("Расход материала: ", "Расход"), ("Цена материала: ", "Цена_мат"), ("Цена работы (ед.): ", "Цена_раб")]
        for frame, suffix in ((frame_v1, "_1"), (frame_v2, "_2")):
            for r, (label, base) in enumerate(variant_fields):
                key = base + suffix
                tk.Label(frame, text=label).grid(row=r, column=0, sticky="w", pady=2)
                en = tk.Entry(frame, width=16)
                add_clipboard_support(en)
                add_context_menu(en)
                en.grid(row=r, column=1, padx=5, pady=2)
                self.entries[key] = en
        
        # ✅ ЧЕКБОКС "БЕЗ МАТЕРИАЛОВ"
        self.no_materials_var = tk.BooleanVar(value=False)
        chk_frame = tk.Frame(frame_input)
        chk_frame.grid(row=4, column=0, columnspan=4, pady=5, sticky="w")
        tk.Checkbutton(chk_frame, text="Без материалов (только работа)", 
                       variable=self.no_materials_var, font=("Arial", 10)).pack(side=tk.LEFT)
        tk.Label(chk_frame, text="  (если отмечено — поля материала игнорируются)", 
                 fg="gray", font=("Arial", 9)).pack(side=tk.LEFT)

        btn_f = tk.Frame(self.tab_db)
        btn_f.pack(pady=5)
        tk.Button(btn_f, text="Сохранить в базу", bg="#4CAF50", fg="white",
                  command=self.save_to_db).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_f, text="Удалить из базы", bg="#f44336", fg="white",
                  command=self.delete_from_db).pack(side=tk.LEFT, padx=5)

        filter_frame = tk.Frame(self.tab_db)
        filter_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(filter_frame, text="🔍 Поиск работы:").pack(side=tk.LEFT)
        self.filter_work = tk.Entry(filter_frame, width=35)
        self.filter_work.pack(side=tk.LEFT, padx=5)
        tk.Label(filter_frame, text="🔍 Поиск материала:").pack(side=tk.LEFT, padx=(10, 0))
        self.filter_mat = tk.Entry(filter_frame, width=35)
        self.filter_mat.pack(side=tk.LEFT, padx=5)
        tk.Button(filter_frame, text="Сброс", command=self.clear_db_filters).pack(side=tk.RIGHT, padx=5)
        self.filter_work.bind('<KeyRelease>', lambda e: self.refresh_db_table())
        self.filter_mat.bind('<KeyRelease>', lambda e: self.refresh_db_table())

        db_tree_frame = tk.Frame(self.tab_db)
        db_tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree_db = ttk.Treeview(db_tree_frame, columns=self.display_cols, show='headings')
        for c in self.display_cols:
            self.tree_db.heading(c, text=c, command=lambda _col=c: self.sort_column(_col))
            width = 250 if c == "Работа" else 100 if "Цена" in c else 80
            self.tree_db.column(c, width=width)
        db_yscroll = ttk.Scrollbar(db_tree_frame, orient="vertical", command=self.tree_db.yview)
        self.tree_db.configure(yscrollcommand=db_yscroll.set)
        self.tree_db.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        db_yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_db.bind("<<TreeviewSelect>>", self.load_to_entries)
        self.refresh_db_table()

    def clear_db_filters(self):
        self.filter_work.delete(0, tk.END)
        self.filter_mat.delete(0, tk.END)
        self.refresh_db_table()

    def sort_column(self, col):
        self.sort_orders[col] = not self.sort_orders[col]
        self.db = self.db.sort_values(by=col, ascending=self.sort_orders[col]).reset_index(drop=True)
        self.refresh_db_table()

    def save_to_db(self):
        """Сохраняет запись в базу данных."""
        try:
            data = {'Работа': self.work_text.get("1.0", tk.END).strip()}
            for c in ('Ед_изм_раб', 'Материал', 'Ед_изм'):
                data[c] = self.entries[c].get().strip()
            
            if not data['Работа']:
                return messagebox.showerror("Ошибка", "Введите наименование работы.")
            
            # ✅ Проверяем чекбокс "Без материалов"
            no_materials = self.no_materials_var.get()
            if no_materials or data['Материал'] in ("", "-", "0"):
                data['Материал'] = "-"
                data['Ед_изм'] = "-"
                for nf in ('Расход_1', 'Цена_мат_1', 'Расход_2', 'Цена_мат_2'):
                    data[nf] = 0.0
            else:
                for nf in ('Расход_1', 'Цена_мат_1', 'Расход_2', 'Цена_мат_2'):
                    raw = self.entries[nf].get().strip()
                    if raw in ("", "-", " "):
                        data[nf] = 0.0
                    else:
                        try:
                            data[nf] = float(raw.replace(',', '.'))
                        except ValueError:
                            return messagebox.showerror("Ошибка", f"Поле «{nf}» должно быть числом.")
            
            # ✅ Парсим цены работы (с автоподстановкой В1 от В2)
            mat_ratio = self.settings.get('auto_price_mat_ratio', 1.35)
            work_ratio = self.settings.get('auto_price_work_ratio', 2.2)
            
            for price_field in ('Цена_раб_1', 'Цена_раб_2'):
                raw = self.entries[price_field].get().strip()
                if raw == "" or raw == "-":
                    data[price_field] = 0.0
                else:
                    try:
                        data[price_field] = float(raw.replace(',', '.'))
                    except ValueError:
                        return messagebox.showerror("Ошибка", f"Поле «{price_field}» должно быть числом.")
            
            # ✅ Автоподстановка В1 от В2, если В1 не задан
            if data['Цена_раб_2'] > 0 and data['Цена_раб_1'] == 0.0:
                data['Цена_раб_1'] = round(data['Цена_раб_2'] * work_ratio, 2)
                self.entries['Цена_раб_1'].delete(0, tk.END)
                self.entries['Цена_раб_1'].insert(0, str(data['Цена_раб_1']))
            
            if not no_materials and data['Материал'] != "-":
                if data['Цена_мат_2'] > 0 and data['Цена_мат_1'] == 0.0:
                    data['Цена_мат_1'] = round(data['Цена_мат_2'] * mat_ratio, 2)
                    self.entries['Цена_мат_1'].delete(0, tk.END)
                    self.entries['Цена_мат_1'].insert(0, str(data['Цена_мат_1']))
            
            work_name = data['Работа']
            
            if self.db_manager is not None:
                # ✅ Сохранение через нормализованную БД (Parquet)
                work = self.db_manager.get_work_by_name(work_name)
                if work is None:
                    work_id = self.db_manager.add_work(
                        work_name, data['Ед_изм_раб'],
                        data['Цена_раб_1'], data['Цена_раб_2']
                    )
                else:
                    work_id = work['id']
                    self.db_manager.update_work(
                        work_id, unit=data['Ед_изм_раб'],
                        price_1=data['Цена_раб_1'], price_2=data['Цена_раб_2']
                    )
                
                # ✅ Очищаем старые связи перед добавлением новых
                self.db_manager.delete_work_material_links_by_work(work_id)
                
                mat_name = data['Материал']
                if mat_name != "-":
                    mat = self.db_manager.get_material_by_name(mat_name)
                    if mat is None:
                        mat_id = self.db_manager.add_material(
                            mat_name, data['Ед_изм'],
                            data['Цена_мат_1'], data['Цена_мат_2']
                        )
                    else:
                        mat_id = mat['id']
                        self.db_manager.update_material(
                            mat_id, unit=data['Ед_изм'],
                            price_1=data['Цена_мат_1'], price_2=data['Цена_мат_2']
                        )
                    self.db_manager.add_work_material_link(
                        work_id, mat_id,
                        data['Расход_1'], data['Расход_2']
                    )
                
                self.db_manager.flush()
                self.db = self.db_manager.get_legacy_dataframe()
            else:
                # ✅ Старый метод с Excel
                mask = ((self.db['Работа'].astype(str).str.strip() == work_name)
                        & (self.db['Материал'].astype(str).str.strip() == data['Материал']))
                self.db = self.db[~mask].reset_index(drop=True)
                new_row = pd.DataFrame([data], columns=sc.COLS)
                
                self.db = pd.concat([self.db, new_row], ignore_index=True)
                self.db = pd.concat([self.db, pd.DataFrame([data])[sc.COLS]], ignore_index=True)
                os.makedirs(self.db_folder, exist_ok=True)
                self.db.to_excel(self.db_file, index=False)
            
            if hasattr(self, 'filter_work'):
                self.filter_work.delete(0, tk.END)
            if hasattr(self, 'filter_mat'):
                self.filter_mat.delete(0, tk.END)
            
            self.refresh_db_table()
            self.update_combobox()
            
            msg = "Запись сохранена в базе."
            if no_materials:
                msg += "\n(работа без материалов)"
            messagebox.showinfo("Готово", msg)
            self.refresh_estimate_if_needed(work_name)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")

    def refresh_estimate_if_needed(self, work_name):
        rows = self._gather_rows()
        used = any(sc.is_work(str(v[1])) and sc.clean_name(v[1]) == work_name for v in rows)
        if not used:
            return
        new_rows = []
        i, n = 0, len(rows)
        while i < n:
            vals = rows[i]
            name = str(vals[1]).strip()
            if sc.is_work(name) and sc.clean_name(name) == work_name:
                vol = sc.to_float(vals[4], 1.0)
                block = sc.build_work_block(work_name, vol, self.db, 1, self.db_manager)
                if block:
                    new_rows.extend(block)
                i += 1
                while i < n and (sc.is_material(str(rows[i][1])) or sc.is_total(str(rows[i][1]))):
                    i += 1
                continue
            new_rows.append(vals)
            i += 1
        self.tree_smeta.delete(*self.tree_smeta.get_children())
        for vals in new_rows:
            tags = ("section",) if sc.is_section(str(vals[1])) else ()
            self.tree_smeta.insert("", tk.END, values=vals, tags=tags)
        self.full_rebuild()

    def delete_from_db(self):
        sel = self.tree_db.selection()
        if not sel:
            return messagebox.showwarning("Внимание", "Выберите строку для удаления.")
        if not messagebox.askyesno("Подтверждение", "Удалить выбранную запись из базы?"):
            return
        try:
            idx_to_delete = int(sel[0])
            if idx_to_delete not in self.db.index:
                return messagebox.showerror("Ошибка", "Индекс строки не найден в базе.")
            
            work_name = str(self.db.loc[idx_to_delete, 'Работа']).strip()
            
            if self.db_manager is not None:
                work = self.db_manager.get_work_by_name(work_name)
                if work is None:
                    return messagebox.showerror("Ошибка", f"Работа «{work_name}» не найдена в базе.")
                work_id = work['id']
                self.db_manager.delete_work(work_id)
                self.db_manager.flush()
                self.db = self.db_manager.get_legacy_dataframe()
            else:
                mask = self.db['Работа'].astype(str).str.strip() == work_name
                self.db = self.db[~mask].reset_index(drop=True)
                os.makedirs(self.db_folder, exist_ok=True)
                self.db.to_excel(self.db_file, index=False)
            
            self.refresh_db_table()
            self.update_combobox()
            messagebox.showinfo("Готово", f"Запись «{work_name}» удалена.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось удалить запись:\n{e}")

    def refresh_db_table(self):
        if not hasattr(self, 'tree_db'):
            return
        self.tree_db.delete(*self.tree_db.get_children())
        df = self.db.copy()
        work_q = self.filter_work.get().strip().lower()
        mat_q = self.filter_mat.get().strip().lower()
        if work_q:
            df = df[df['Работа'].astype(str).str.lower().str.contains(work_q, na=False)]
        if mat_q:
            df = df[df['Материал'].astype(str).str.lower().str.contains(mat_q, na=False)]
        for idx, r in df.iterrows():
            vals = [r[c] for c in self.display_cols]
            self.tree_db.insert("", tk.END, iid=str(idx), values=vals)

    def load_to_entries(self, event):
        sel = self.tree_db.selection()
        if not sel:
            return
        try:
            idx = int(sel[0])
            row = self.db.loc[idx]
        except (KeyError, ValueError):
            return
        self.work_text.delete("1.0", tk.END)
        self.work_text.insert("1.0", row['Работа'])
        for c in sc.COLS[1:]:
            self.entries[c].delete(0, tk.END)
            self.entries[c].insert(0, row[c])
        # ✅ Сбрасываем чекбокс "Без материалов" в зависимости от содержимого
        mat_val = str(row.get('Материал', '')).strip()
        self.no_materials_var.set(mat_val in ("", "-", "0"))

    def update_combobox(self):
        if not hasattr(self, 'work_combo'):
            return
        if not self.db.empty:
            self.all_works_list = sorted(list(self.db['Работа'].astype(str).unique()))
            self.work_combo['values'] = self.all_works_list
        else:
            self.all_works_list = []
            self.work_combo['values'] = []

    def filter_works_combo(self, event=None):
        typed = self.work_combo.get().strip()
        if not hasattr(self, 'all_works_list'):
            return
        if not typed:
            self.work_combo['values'] = self.all_works_list
            return
        lower_typed = typed.lower()
        filtered = [w for w in self.all_works_list if lower_typed in w.lower()]
        self.work_combo['values'] = filtered
        self.work_combo.icursor(tk.END)

    def refresh_db_list(self):
        if not os.path.exists(self.db_folder):
            os.makedirs(self.db_folder, exist_ok=True)
        files = os.listdir(self.db_folder)
        
        # Собираем Excel файлы
        db_files = [f for f in files if f.lower().endswith('.xlsx')]
        
        # Собираем Parquet базы (группы из 3 файлов)
        if HAS_DB_MANAGER:
            parquet_bases = set()
            for f in files:
                if f.lower().endswith('_works.parquet'):
                    base_name = f[:-len('_works.parquet')]
                    if (os.path.exists(os.path.join(self.db_folder, f'{base_name}_works.parquet')) and
                        os.path.exists(os.path.join(self.db_folder, f'{base_name}_materials.parquet')) and
                        os.path.exists(os.path.join(self.db_folder, f'{base_name}_work_materials.parquet'))):
                        parquet_bases.add(f'{base_name}.parquet')
            db_files.extend(sorted(parquet_bases))
        
        db_files = sorted(db_files)
        self.db_combo['values'] = db_files
        
        if self.active_db_filename in db_files:
            self.db_combo.set(self.active_db_filename)
        elif db_files:
            self.db_combo.set(db_files[0])
            self.on_db_selected()

    def on_db_selected(self, event=None):
        new_db = self.db_combo.get()
        if new_db and new_db != self.active_db_filename:
            self.active_db_filename = new_db
            
            # Для Parquet баз
            if new_db.endswith('.parquet') and HAS_DB_MANAGER:
                db_name = new_db[:-len('.parquet')]
                self.db_manager = DatabaseManager(self.db_folder, db_name)
                self.db_file = os.path.join(self.db_folder, f'{db_name}_works.parquet')
            else:
                # Excel база
                if HAS_DB_MANAGER:
                    self.db_manager = None
                self.db_file = os.path.join(self.db_folder, new_db)
            
            self._save_active_db_to_settings()
            self.db = self._load_db()
            self.refresh_db_table()
            self.update_combobox()

    def _save_active_db_to_settings(self):
        settings = load_settings()
        settings['active_db_filename'] = self.active_db_filename
        save_settings(settings)

    def create_new_db(self):
        new_name = simpledialog.askstring("Новая база", "Введите имя файла (например, База_Кровля.xlsx):")
        if not new_name:
            return
        new_name = re.sub(r'[<>:"/\\|?*]', '_', new_name)
        if not new_name.lower().endswith('.xlsx'):
            new_name += '.xlsx'
        new_path = os.path.join(self.db_folder, new_name)
        if os.path.exists(new_path):
            return messagebox.showwarning("Внимание", "Файл уже существует.")
        
        if self.db_manager:
            db_name = new_name.replace('.xlsx', '')
            self.active_db_filename = db_name
            self.db_manager = DatabaseManager(self.db_folder, db_name)
        else:
            pd.DataFrame(columns=sc.COLS).to_excel(new_path, index=False)
        
        self.refresh_db_list()
        self.db_combo.set(new_name)
        self.on_db_selected()
        messagebox.showinfo("Готово", f"База «{new_name}» создана и загружена.")

    def delete_current_db(self):
        if not self.active_db_filename:
            return
        if not messagebox.askyesno("Подтверждение", f"Удалить базу «{self.active_db_filename}»?\nЭто действие нельзя отменить."):
            return
        path = os.path.join(self.db_folder, self.active_db_filename)
        try:
            os.remove(path)
            if self.db_manager:
                self.db_manager = None
            self.refresh_db_list()
            if self.db_combo['values']:
                self.db_combo.set(self.db_combo['values'][0])
                self.on_db_selected()
            else:
                self.db = pd.DataFrame(columns=sc.COLS)
                self.refresh_db_table()
                self.update_combobox()
            messagebox.showinfo("Готово", "База удалена.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось удалить файл:\n{e}")

    def setup_calc_tab(self):
        frame_title = tk.Frame(self.tab_calc, pady=4)
        frame_title.pack(fill=tk.X, padx=10)
        tk.Label(frame_title, text="Наименование сметы / объект:").pack(side=tk.LEFT)
        self.title_entry = tk.Entry(frame_title)
        add_clipboard_support(self.title_entry)
        add_context_menu(self.title_entry)
        self.title_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        frame_top = tk.Frame(self.tab_calc, pady=5)
        frame_top.pack(fill=tk.X, padx=10)
        self.work_combo = ttk.Combobox(frame_top, width=120)
        self.work_combo.pack(side=tk.LEFT, padx=5)
        self.work_combo.bind('<KeyRelease>', self.filter_works_combo)
        self.update_combobox()
        tk.Label(frame_top, text="Объём:").pack(side=tk.LEFT)
        self.vol_entry = tk.Entry(frame_top, width=10)
        add_clipboard_support(self.vol_entry)
        add_context_menu(self.vol_entry)
        self.vol_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(frame_top, text="Добавить работу", bg="#2196F3", fg="white",
                  command=self.add_to_estimate).pack(side=tk.LEFT, padx=10)
        frame_tools = tk.Frame(self.tab_calc, pady=5)
        frame_tools.pack(fill=tk.X, padx=10)
        tk.Button(frame_tools, text="➕ Раздел", bg="#9C27B0", fg="white",
                  command=self.add_section).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_tools, text="➕ Материал", bg="#4CAF50", fg="white",
                  command=self.add_material_to_selected).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_tools, text="📋 Дубль мат.", bg="#FF9800", fg="white",
                  command=self.duplicate_material).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_tools, text=" Удалить", bg="#f44336", fg="white",
                  command=self.remove_smeta_row).pack(side=tk.LEFT, padx=5)
        self.calc_cols = sc.CALC_HEADERS

        calc_tree_frame = tk.Frame(self.tab_calc)
        calc_tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        style = ttk.Style()
        style.configure("Smeta.Treeview", rowheight=30)
        self.tree_smeta = ttk.Treeview(
            calc_tree_frame, 
            columns=self.calc_cols, 
            show='headings', 
            selectmode='extended',
            style="Smeta.Treeview"
        )
        widths = {"№": 45, "Наименование": 380, "Ед. изм.": 65}
        for c in self.calc_cols:
            self.tree_smeta.heading(c, text=c)
            self.tree_smeta.column(c, width=widths.get(c, 90),
                                    anchor="w" if c == "Наименование" else "center")
        calc_yscroll = ttk.Scrollbar(calc_tree_frame, orient="vertical", command=self.tree_smeta.yview)
        self.tree_smeta.configure(yscrollcommand=calc_yscroll.set)
        self.tree_smeta.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        calc_yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_smeta.tag_configure("section", background="#E1BEE7", font=("Arial", 10, "bold"))
        self.tree_smeta.bind("<Double-1>", self.on_tree_double_click)
       
        def handle_undo(event):
            if event.keycode == 90 and (event.state & 0x4 or event.state & 0x10):
                self.undo_action(event)
                return "break"

        self.tree_smeta.bind("<Control-Key>", handle_undo, add="+")
        self.tree_smeta.bind("<Command-Key>", handle_undo, add="+")
        self.tree_smeta.bind("<Button-3>", self._show_context_menu)
        self._tooltip_win = None
        self.tree_smeta.bind("<Motion>", self._show_name_tooltip)
        self.tree_smeta.bind("<Leave>", self._hide_tooltip)
        frame_bottom = tk.Frame(self.tab_calc, pady=5)
        frame_bottom.pack(fill=tk.X, padx=10)
        tk.Button(frame_bottom, text="📂 Загрузить смету", bg="#607D8B", fg="white",
                  command=self.load_estimate).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_bottom, text="Выгрузить в Excel", bg="#FF9800", fg="white",
                  command=self.export_excel).pack(side=tk.LEFT, padx=5)
        frame_extra = tk.LabelFrame(self.tab_calc, text="Доп. расходы (учитываются в ИТОГО)", padx=8, pady=4)
        frame_extra.pack(fill=tk.X, padx=10, pady=3)
        self.extra_entries = {}
        specs = [
            ("Накладные/транспортные, В1: ", "overhead1"), ("В2: ", "overhead2"),
            ("   Подъёмные механизмы, В1: ", "lift1"), ("В2: ", "lift2"),
            ("   Вывоз мусора, В1: ", "trash1"), ("В2: ", "trash2"),
        ]
        for label, key in specs:
            tk.Label(frame_extra, text=label).pack(side=tk.LEFT, padx=(4, 2))
            en = tk.Entry(frame_extra, width=12)
            add_clipboard_support(en)
            add_context_menu(en)
            en.insert(0, "0")
            en.pack(side=tk.LEFT, padx=2)
            en.bind('<KeyRelease>', lambda e: self.update_total_sum())
            self.extra_entries[key] = en
        total_frame = tk.Frame(self.tab_calc)
        total_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        tk.Label(total_frame, text=" ").pack(side=tk.LEFT, expand=True)
        self.total_label = tk.Label(total_frame, text="ИТОГО В1: 0.00 ₽   |   ИТОГО В2: 0.00 ₽   |   Экономия: 0.00 ₽ (0.0%)", font=("Arial", 13, "bold"), fg="#1565C0")
        self.total_label.pack(side=tk.RIGHT)

    def _gather_rows(self):
        return [self.tree_smeta.item(it, 'values') for it in self.tree_smeta.get_children()]

    def full_rebuild(self):
        rows = self._gather_rows()
        rebuilt = sc.rebuild_smeta(rows)
        self.tree_smeta.delete(*self.tree_smeta.get_children())
        for vals in rebuilt:
            tags = ("section",) if sc.is_section(str(vals[1])) else ()
            self.tree_smeta.insert("", tk.END, values=vals, tags=tags)
        self.update_total_sum()

    def add_section(self):
        self.tree_smeta.insert("", tk.END, values=("", "РАЗДЕЛ: Новый раздел", "", "", "", "", "", "", "", "", ""), tags=("section",))

    def _editable_cols_for(self, name_raw):
        name_raw = str(name_raw).strip()
        if sc.is_work(name_raw):
            return {1, 4, 5, 9}
        if sc.is_material(name_raw):
            return {1, 3, 5, 7, 9}
        return set()

    def on_tree_double_click(self, event):
        item = self.tree_smeta.identify_row(event.y)
        if not item:
            return
        col = self.tree_smeta.identify_column(event.x)
        if not col or col == '#0':
            return
        col_idx = int(col[1:]) - 1
        vals = self.tree_smeta.item(item, 'values')
        if not vals:
            return
        name_raw = str(vals[1]).strip()
        if sc.is_section(name_raw) and col_idx == 1:
            self._edit_section_inline(item, vals)
            return
        if col_idx in self._editable_cols_for(name_raw):
            self._start_cell_edit(item, vals, col_idx)

    def add_to_estimate(self):
        work_name = self.work_combo.get().strip()
        if not work_name:
            return messagebox.showwarning("Внимание", "Выберите работу из списка.")
        raw = self.vol_entry.get().strip()
        try:
            vol = float(raw.replace(',', '.'))
        except ValueError:
            return messagebox.showerror("Ошибка", "Введите корректное число в поле «Объём».")
        self._add_work_to_smeta(work_name, vol)

    def _add_work_to_smeta(self, work_name, volume, suppress_total_update=False):
        rows = self._gather_rows()
        next_num = 1
        for vals in rows:
            if vals and str(vals[0]).strip().isdigit():
                next_num = max(next_num, int(vals[0]) + 1)
        block = sc.build_work_block(work_name, volume, self.db, next_num, self.db_manager)
        if block is None:
            return messagebox.showwarning("Внимание", f"Работа «{work_name}» не найдена в базе!")
        for vals in block:
            self.tree_smeta.insert("", tk.END, values=vals)
        if not suppress_total_update:
            self.update_total_sum()

    def remove_smeta_row(self):
        selected = self.tree_smeta.selection()
        if not selected:
            return
        for item in selected:
            vals = self.tree_smeta.item(item, 'values')
            if vals and sc.is_total(str(vals[1])):
                return messagebox.showwarning("Внимание", "Нельзя удалить итоговую строку напрямую.\nДля удаления узла целиком удалите строку с работой.")
        for item in selected:
            self.tree_smeta.delete(item)
        self.full_rebuild()

    def update_total_sum(self):
        rows = self._gather_rows()
        t1, t2 = sc.compute_grand_totals(rows)
        oh1 = sc.to_float(self.extra_entries.get('overhead1', {}).get())
        oh2 = sc.to_float(self.extra_entries.get('overhead2', {}).get())
        l1 = sc.to_float(self.extra_entries.get('lift1', {}).get())
        l2 = sc.to_float(self.extra_entries.get('lift2', {}).get())
        tr1 = sc.to_float(self.extra_entries.get('trash1', {}).get())
        tr2 = sc.to_float(self.extra_entries.get('trash2', {}).get())
        grand1 = round(t1 + oh1 + l1 + tr1, 2)
        grand2 = round(t2 + oh2 + l2 + tr2, 2)
        saving = round(grand1 - grand2, 2)
        pct = (saving / grand1 * 100.0) if grand1 else 0.0
        def fmt(v):
            return format(v, ',.2f').replace(',', ' ')
        self.total_label.config(
            text=f"ИТОГО В1: {fmt(grand1)} ₽   |   ИТОГО В2: {fmt(grand2)} ₽   |   "
                 f"Экономия: {fmt(saving)} ₽ ({pct:.1f}%)"
        )
        self.root.update_idletasks()

    def _edit_section_inline(self, item, vals):
        bbox = self.tree_smeta.bbox(item, column='#2')
        if not bbox:
            return
        x, y, w, h = bbox
        root_x = self.tree_smeta.winfo_rootx() - self.root.winfo_rootx()
        root_y = self.tree_smeta.winfo_rooty() - self.root.winfo_rooty()
        self.edit_entry = tk.Entry(self.root, font=("Arial", 10), bd=1, relief="solid")
        self.edit_entry.place(x=root_x + x, y=root_y + y, width=w, height=h)
        add_clipboard_support(self.edit_entry)
        add_context_menu(self.edit_entry)
        self.edit_entry.insert(0, vals[1].replace("РАЗДЕЛ: ", ""))
        self.edit_entry.focus()
        self.edit_entry.select_range(0, tk.END)
        self.edit_item, self.edit_orig, self.edit_col_idx = item, list(vals), 1
        self.edit_entry.bind("<Return>", lambda e: self._save_section_edit())
        self.edit_entry.bind("<Escape>", lambda e: self._destroy_edit())
        self.edit_entry.bind("<FocusOut>", lambda e: self.root.after(100, self._save_section_edit))

    def _save_section_edit(self):
        if self.edit_entry is None or not self.edit_entry.winfo_exists():
            return
        nm = self.edit_entry.get().strip()
        if nm:
            nv = self.edit_orig[:]
            nv[1] = f"РАЗДЕЛ: {nm}"
            self.tree_smeta.item(self.edit_item, values=tuple(nv))
        self._destroy_edit()

    def _start_cell_edit(self, item, vals, col_idx):
        self.undo_stack.append((item, col_idx, vals[col_idx]))
        ct = f'#{col_idx + 1}'
        bbox = self.tree_smeta.bbox(item, column=ct)
        if not bbox:
            return
        x, y, w, h = bbox
        root_x = self.tree_smeta.winfo_rootx() - self.root.winfo_rootx()
        root_y = self.tree_smeta.winfo_rooty() - self.root.winfo_rooty()
        self.edit_entry = tk.Entry(self.root, font=("Arial", 10), bd=1, relief="solid")
        self.edit_entry.place(x=root_x + x, y=root_y + y, width=w, height=h)
        add_clipboard_support(self.edit_entry)
        add_context_menu(self.edit_entry)
        self.edit_entry.insert(0, str(vals[col_idx]))
        self.edit_entry.focus()
        self.edit_entry.select_range(0, tk.END)
        self.edit_item, self.edit_orig, self.edit_col_idx = item, list(vals), col_idx
        self.edit_entry.bind("<Return>", lambda e: self._finish_cell_edit())
        self.edit_entry.bind("<Escape>", lambda e: self._destroy_edit())
        self.edit_entry.bind("<FocusOut>", lambda e: self.root.after(100, self._finish_cell_edit))

    def _finish_cell_edit(self):
        if self.edit_entry is None or not self.edit_entry.winfo_exists():
            return
        nv = self.edit_entry.get().strip()
        if nv == "":
            return self._destroy_edit()
        nvals = self.edit_orig[:]
        numeric_cols = {3, 4, 5, 7, 9}
        if self.edit_col_idx in numeric_cols:
            try:
                nvals[self.edit_col_idx] = float(nv.replace(',', '.'))
            except ValueError:
                messagebox.showerror("Ошибка", "Введите корректное число.")
                self.edit_entry.focus()
                return
        else:
            nvals[self.edit_col_idx] = nv
        if self.edit_col_idx == 1:
            orig = str(self.edit_orig[1]).strip()
            clean = str(nvals[1]).replace("Работа: ", "").replace("    > ", "").strip()
            if orig.startswith("    > "):
                nvals[1] = f"    > {clean}"
            elif orig.startswith("Работа: "):
                nvals[1] = f"Работа: {clean}"
        self.tree_smeta.item(self.edit_item, values=tuple(nvals))
        self._sync_db(nvals, self.edit_col_idx)
        self._propagate_same_name(nvals, self.edit_col_idx)
        try:
            self.full_rebuild()
        except Exception as e:
            messagebox.showerror("Ошибка пересчёта", f"Не удалось обновить итоги:\n{e}")
        finally:
            self._destroy_edit()

    def _propagate_same_name(self, nvals, col_idx):
        if col_idx not in (3, 5, 7, 9):
            return
        orig_name = str(self.edit_orig[1]).strip()
        is_work_row = orig_name.startswith("Работа: ")
        is_mat_row = orig_name.startswith("    > ")
        if not (is_work_row or is_mat_row):
            return
        clean = sc.clean_name(orig_name)
        new_val = nvals[col_idx]
        for item in self.tree_smeta.get_children():
            if item == self.edit_item:
                continue
            vals = list(self.tree_smeta.item(item, 'values'))
            v_name = str(vals[1]).strip()
            same_type = ((is_work_row and v_name.startswith("Работа: ")) or (is_mat_row and v_name.startswith("    > ")))
            if not same_type:
                continue
            if sc.clean_name(v_name) == clean:
                vals[col_idx] = new_val
                self.tree_smeta.item(item, values=tuple(vals))

    def _sync_db(self, nv, col_idx):
        ov = str(self.edit_orig[1]).strip()
        is_work_row = ov.startswith("Работа: ")
        is_mat_row = ov.startswith("    > ")
        clean = sc.clean_name(ov)
        col_map_mat = {3: 'consumption_1', 5: 'price_1', 7: 'consumption_2', 9: 'price_2'}
        col_map_work = {5: 'price_1', 9: 'price_2'}
        try:
            if self.db_manager is None:
                return
            if col_idx == 1:
                nc = str(nv[1]).replace("Работа: ", "").replace("    > ", "").strip()
                if is_work_row:
                    self.db_manager.works_cache.loc[self.db_manager.works_cache['name'].str.strip() == clean, 'name'] = nc
                elif is_mat_row:
                    self.db_manager.materials_cache.loc[self.db_manager.materials_cache['name'].str.strip() == clean, 'name'] = nc
                self.db_manager.works_dirty = True
                self.db_manager.materials_dirty = True
            elif is_mat_row and col_idx in col_map_mat:
                mat = self.db_manager.get_material_by_name(clean)
                if mat is not None:
                    self.db_manager.work_materials_cache.loc[self.db_manager.work_materials_cache['material_id'] == mat['id'], col_map_mat[col_idx]] = sc.to_float(nv[col_idx])
                    self.db_manager.work_materials_dirty = True
            elif is_work_row and col_idx in col_map_work:
                self.db_manager.works_cache.loc[self.db_manager.works_cache['name'].str.strip() == clean, col_map_work[col_idx]] = sc.to_float(nv[col_idx])
                self.db_manager.works_dirty = True
            else:
                return
            self.db_manager.flush()
            self.refresh_db_table()
            self.update_combobox()
        except Exception as e:
            messagebox.showerror("Ошибка БД", str(e))

    def _destroy_edit(self):
        if self.edit_entry is not None and self.edit_entry.winfo_exists():
            self.edit_entry.destroy()
        self.edit_entry = None
        self.edit_item = None
        self.edit_col_idx = None
        self.edit_orig = None

    def _show_name_tooltip(self, event):
        item = self.tree_smeta.identify_row(event.y)
        col = self.tree_smeta.identify_column(event.x)
        if not item or col != '#2':
            self._hide_tooltip()
            return
        vals = self.tree_smeta.item(item, 'values')
        if not vals:
            self._hide_tooltip()
            return
        text = str(vals[1])
        if len(text) <= 45:
            self._hide_tooltip()
            return
        self._hide_tooltip()
        self._tooltip_win = tk.Toplevel(self.root)
        self._tooltip_win.wm_overrideredirect(True)
        self._tooltip_win.wm_geometry(f"+{event.x_root+15}+{event.y_root+15}")
        self._tooltip_win.attributes("-topmost", True)
        tk.Label(
            self._tooltip_win, 
            text=text, 
            background="#ffffdd", 
            relief="solid", 
            borderwidth=1,
            wraplength=500, 
            justify="left", 
            font=("Arial", 11),
            padx=8, pady=4
        ).pack()

    def _hide_tooltip(self, event=None):
        if hasattr(self, '_tooltip_win') and self._tooltip_win is not None:
            try:
                self._tooltip_win.destroy()
            except tk.TclError:
                pass
            finally:
                self._tooltip_win = None

    def _open_material_dialog(self, initial_data=None):
        win = tk.Toplevel(self.root)
        win.title("Добавить материал")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        result = {}
        entries = {}
        fields = [("Наименование: ", "name", 30), ("Ед. изм.: ", "unit", 12), ("Норма расхода (В1): ", "norm1", 14), ("Цена за ед. (В1): ", "price1", 14), ("Норма расхода (В2): ", "norm2", 14), ("Цена за ед. (В2): ", "price2", 14)]
        for i, (lbl, key, w) in enumerate(fields):
            tk.Label(win, text=lbl).grid(row=i, column=0, padx=5, pady=4, sticky="e")
            e = tk.Entry(win, width=w)
            add_clipboard_support(e)
            add_context_menu(e)
            if initial_data and key in initial_data:
                e.insert(0, str(initial_data[key]))
            e.grid(row=i, column=1, padx=5, pady=4, sticky="w")
            entries[key] = e
        err_label = tk.Label(win, text=" ", fg="red")
        err_label.grid(row=len(fields), column=0, columnspan=2)
        def apply():
            name = entries['name'].get().strip()
            if not name:
                err_label.config(text="Введите наименование материала.")
                return
            data = {'name': name, 'unit': entries['unit'].get().strip()}
            for key in ('norm1', 'price1', 'norm2', 'price2'):
                raw = entries[key].get().strip()
                if raw == "" or raw == "-":
                    data[key] = 0.0
                    continue
                try:
                    data[key] = float(raw.replace(',', '.'))
                except ValueError:
                    err_label.config(text=f"Поле «{key}» должно быть числом.")
                    return
            result.update(data)
            win.destroy()
        def cancel():
            result.clear()
            win.destroy()
        btn_f = tk.Frame(win)
        btn_f.grid(row=len(fields) + 1, column=0, columnspan=2, pady=10)
        tk.Button(btn_f, text="Добавить", bg="#4CAF50", fg="white", command=apply).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_f, text="Отмена", command=cancel).pack(side=tk.LEFT, padx=5)
        win.bind("<Return>", lambda e: apply())
        win.wait_window()
        return result if result else None

    def add_material_to_selected(self):
        sel = self.tree_smeta.selection()
        if not sel:
            return messagebox.showwarning("Внимание", "Выберите работу или материал.")
        item = sel[0]
        vals = self.tree_smeta.item(item, 'values')
        name_raw = str(vals[1]).strip()
        if sc.is_section(name_raw):
            return messagebox.showwarning("Внимание", "Нельзя добавить материал в раздел. Выберите работу.")
        if sc.is_total(name_raw):
            return messagebox.showwarning("Внимание", "Выберите строку работы или материала, не итоговую строку.")
        children = list(self.tree_smeta.get_children())
        insert_idx = children.index(item) + 1
        data = self._open_material_dialog()
        if not data or not data.get('name'):
            return
        self.tree_smeta.insert("", insert_idx, values=("", f"    > {data['name']}", data['unit'], data['norm1'], "", data['price1'], "", data['norm2'], "", data['price2'], ""))
        self.full_rebuild()

    def duplicate_material(self):
        sel = self.tree_smeta.selection()
        if not sel:
            return messagebox.showwarning("Внимание", "Выберите строку материала.")
        item = sel[0]
        vals = list(self.tree_smeta.item(item, 'values'))
        name_raw = str(vals[1]).strip()
        if not sc.is_material(name_raw):
            return messagebox.showwarning("Внимание", "Выбрана не строка материала.")
        children = list(self.tree_smeta.get_children())
        insert_idx = children.index(item) + 1
        self.tree_smeta.insert("", insert_idx, values=tuple(vals))
        self.full_rebuild()

    def load_estimate(self):
        file_path = filedialog.askopenfilename(title="Выберите файл сметы", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return
        try:
            xl = pd.ExcelFile(file_path)
        except Exception as e:
            return messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")
        meta_df = None
        if "Meta" in xl.sheet_names:
            meta_df = pd.read_excel(xl, sheet_name="Meta")
        else:
            if not messagebox.askyesno("Предупреждение", "В файле отсутствует служебный лист Meta.\nВосстановление сметы может быть неполным.\nПродолжить?"):
                return
        if "Смета" not in xl.sheet_names:
            return messagebox.showerror("Ошибка", "В файле отсутствует лист 'Смета'.")
        smeta_df = pd.read_excel(xl, sheet_name="Смета", header=None)
        if meta_df is not None and not meta_df.empty:
            self._reconcile_meta(meta_df)
        res = sc.parse_exported_sheet(smeta_df.values.tolist())
        self.tree_smeta.delete(*self.tree_smeta.get_children())
        self.title_entry.delete(0, tk.END)
        if res['title']:
            self.title_entry.insert(0, res['title'])
        missing = []
        for entry in res['sequence']:
            if entry[0] == 'section':
                self.tree_smeta.insert("", tk.END, values=("", f"РАЗДЕЛ: {entry[1]}", "", "", "", "", "", "", "", "", ""), tags=("section",))
            else:
                _, name, vol = entry
                clean_name = sc.clean_name(name)
                if (self.db['Работа'].astype(str).str.strip() == clean_name).any():
                    self._add_work_to_smeta(clean_name, vol, suppress_total_update=True)
                else:
                    missing.append(clean_name)
        oh1, oh2 = res['overhead']
        l1, l2 = res['lifting']
        lt1, lt2 = res.get('lifting_trash', (0.0, 0.0))
        for key, val in (('overhead1', oh1), ('overhead2', oh2), 
                         ('lift1', l1), ('lift2', l2), 
                         ('trash1', lt1), ('trash2', lt2)):
            self.extra_entries[key].delete(0, tk.END)
            self.extra_entries[key].insert(0, str(val))
        self.full_rebuild()
        if missing:
            messagebox.showwarning("Внимание", "Не найдены в справочнике и не были восстановлены работы:\n" + "\n".join(missing))
        messagebox.showinfo("Готово", "Смета загружена.")

    def _reconcile_meta(self, meta_df):
        for col in sc.COLS:
            if col not in meta_df.columns:
                meta_df[col] = "-" if col in ('Работа', 'Ед_изм_раб', 'Материал', 'Ед_изм') else 0.0
        numeric_cols = ('Расход_1', 'Цена_мат_1', 'Цена_раб_1', 'Расход_2', 'Цена_мат_2', 'Цена_раб_2')
        new_rows_list = []
        for _, meta_row in meta_df.iterrows():
            work = str(meta_row['Работа']).strip()
            material = str(meta_row['Материал']).strip()
            existing = self.db[(self.db['Работа'].astype(str).str.strip() == work) & (self.db['Материал'].astype(str).str.strip() == material)]
            if existing.empty:
                new_rows_list.append(meta_row)
                continue
            match_found = False
            for _, exist_row in existing.iterrows():
                try:
                    ok = (str(exist_row['Ед_изм_раб']).strip() == str(meta_row['Ед_изм_раб']).strip() and str(exist_row['Ед_изм']).strip() == str(meta_row['Ед_изм']).strip())
                    for nc in numeric_cols:
                        ok = ok and abs(sc.to_float(exist_row[nc]) - sc.to_float(meta_row[nc])) < 1e-6
                    if ok:
                        match_found = True
                        break
                except Exception:
                    continue
            if not match_found:
                new_rows_list.append(meta_row)
        if new_rows_list:
            if messagebox.askyesno("Новые данные", f"В смете найдено {len(new_rows_list)} новых или изменённых записей.\nДобавить их в справочник?"):
                new_rows_df = pd.DataFrame(new_rows_list)
                for nc in numeric_cols:
                    if nc in new_rows_df.columns:
                        new_rows_df[nc] = pd.to_numeric(new_rows_df[nc], errors='coerce').fillna(0.0)
                if self.db_manager:
                    self.db_manager.save_legacy_dataframe(new_rows_df)
                    self.db = self.db_manager.get_legacy_dataframe()
                else:
                    self.db = pd.concat([self.db, new_rows_df[sc.COLS]], ignore_index=True)
                    os.makedirs(self.db_folder, exist_ok=True)
                    self.db.to_excel(self.db_file, index=False)
                self.refresh_db_table()
                self.update_combobox()
                messagebox.showinfo("Готово", "Новые записи добавлены в справочник.")
        else:
            messagebox.showinfo("Информация", "Все записи из сметы уже присутствуют в справочнике.")

    def export_excel(self):
        rows = self._gather_rows()
        if not rows:
            return messagebox.showwarning("Внимание", "Смета пуста — нечего выгружать.")
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            export_path = os.path.join(self.export_folder, f"Smeta_{timestamp}.xlsx")
            os.makedirs(self.export_folder, exist_ok=True)
            seen_works, seen_materials = set(), set()
            for vals in rows:
                name = str(vals[1]).strip()
                if sc.is_work(name):
                    seen_works.add(sc.clean_name(name))
                elif sc.is_material(name):
                    seen_materials.add(sc.clean_name(name))
            meta_rows = []
            if seen_works or seen_materials:
                meta_df = self.db[self.db['Работа'].isin(seen_works) | self.db['Материал'].isin(seen_materials)].drop_duplicates().reset_index(drop=True)
                if not meta_df.empty:
                    meta_rows = meta_df[sc.COLS].values.tolist()
            title = self.title_entry.get().strip()
            oh1 = sc.to_float(self.extra_entries['overhead1'].get())
            oh2 = sc.to_float(self.extra_entries['overhead2'].get())
            l1 = sc.to_float(self.extra_entries['lift1'].get())
            l2 = sc.to_float(self.extra_entries['lift2'].get())
            tr1 = sc.to_float(self.extra_entries['trash1'].get())
            tr2 = sc.to_float(self.extra_entries['trash2'].get())
            sc.export_smeta_to_excel(rows, export_path, title=title, meta_rows=meta_rows, 
                                      overhead1=oh1, overhead2=oh2, lift1=l1, lift2=l2,
                                      trash1=tr1, trash2=tr2)
            messagebox.showinfo("Excel", f"Смета выгружена успешно!\nФайл: {export_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать файл.\nЗакройте Excel и попробуйте снова.\n\n{e}")

    def undo_action(self, event=None):
        if self.undo_stack:
            item_id, col_idx, old_val = self.undo_stack.pop()
            if self.tree_smeta.exists(item_id):
                vals = list(self.tree_smeta.item(item_id, 'values'))
                vals[col_idx] = old_val
                self.tree_smeta.item(item_id, values=tuple(vals))
                self.full_rebuild()

    def _show_context_menu(self, event):
        self.ctx_menu_item = self.tree_smeta.identify_row(event.y)
        self.ctx_menu_col = self.tree_smeta.identify_column(event.x)
        if not self.ctx_menu_item or not self.ctx_menu_col:
            return
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Копировать значение", command=self._copy_cell)
        menu.add_command(label="Вставить значение", command=self._paste_cell)
        menu.tk_popup(event.x_root, event.y_root)

    def _copy_cell(self):
        if self.ctx_menu_item:
            col_idx = int(self.ctx_menu_col[1:]) - 1
            val = self.tree_smeta.item(self.ctx_menu_item, 'values')[col_idx]
            self.root.clipboard_clear()
            self.root.clipboard_append(str(val))

    def _paste_cell(self):
        if not self.ctx_menu_item:
            return
        try:
            clip = self.root.clipboard_get().strip()
        except tk.TclError:
            return
        col_idx = int(self.ctx_menu_col[1:]) - 1
        vals = self.tree_smeta.item(self.ctx_menu_item, 'values')
        name_raw = str(vals[1]).strip()
        if col_idx not in self._editable_cols_for(name_raw):
            return
        self._start_cell_edit(self.ctx_menu_item, vals, col_idx)
        if self.edit_entry and self.edit_entry.winfo_exists():
            self.edit_entry.delete(0, tk.END)
            self.edit_entry.insert(0, clip)
            self.edit_entry.icursor(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = SmetaApp(root)
    root.mainloop()