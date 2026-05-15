import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
from datetime import datetime
import json
from tkinter import simpledialog

# Путь к файлу настроек
SETTINGS_FILE = 'settings.json'
COLS = ['Работа', 'Ед_изм_раб', 'Материал', 'Ед_изм', 'Расход', 'Цена_мат', 'Цена_раб']


def load_settings():
    """Загружает настройки из JSON-файла."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_settings = {
        'db_folder': script_dir,
        'export_folder': script_dir,
        'active_db_filename': 'my_works_base.xlsx'  # 🆕 Имя активной базы
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
    return default_settings


def save_settings(settings):
    """Сохраняет настройки в JSON-файл."""
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Ошибка сохранения настроек: {e}")


def add_clipboard_support(widget):
    """Реализация копирования (Ctrl+C) и вставки (Ctrl+V) для Entry и Text."""
    def copy(event):
        try:
            text = event.widget.selection_get()
            event.widget.clipboard_clear()
            event.widget.clipboard_append(text)
        except tk.TclError:
            pass
        return "break"  # ⛔ Останавливает срабатывание встроенного обработчика Tkinter

    def paste(event):
        try:
            text = event.widget.clipboard_get()
            event.widget.insert(tk.INSERT, text)
        except tk.TclError:
            pass
        return "break"  # ⛔ Останавливает дублирование вставки

    # Убраны пробелы в строках привязки (артефакты копирования)
    widget.bind("<Control-c>", copy)
    widget.bind("<Control-v>", paste)
    widget.bind("<Command-c>", copy)
    widget.bind("<Command-v>", paste)


def add_context_menu(widget):
    """Добавляет контекстное меню (ПКМ) с командами Вырезать, Копировать, Вставить."""
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
            text = widget.clipboard_get()
            widget.insert(tk.INSERT, text)
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


class SmetaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Сметчик PRO 4.5.2")
        self.root.geometry("1450x900")

        # Инициализация путей без global
        settings = load_settings()
        self.db_folder = settings['db_folder']
        self.export_folder = settings['export_folder']
        self.active_db_filename = settings.get('active_db_filename', 'my_works_base.xlsx')
        self.db_file = os.path.join(self.db_folder, self.active_db_filename)

        # Инициализация атрибутов для inline-редактирования
        self.edit_entry = None
        self.edit_item = None
        self.edit_col_idx = None
        self.edit_orig = None
        self.undo_stack = []  # Стек для отмены изменений (Ctrl+Z)
        self.ctx_menu_item = None
        self.ctx_menu_col = None

        self.db = self._load_db()
        self.sort_orders = {col: False for col in COLS}

        # Порядок столбцов в таблице Справочника (визуальный)
        self.display_cols = ['Работа', 'Ед_изм_раб', 'Цена_раб', 'Материал', 'Ед_изм', 'Расход', 'Цена_мат']
        # Меню
        menubar = tk.Menu(root)
        root.config(menu=menubar)
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Настройки", menu=settings_menu)
        settings_menu.add_command(label="Настройки", command=self.open_settings)

        self.notebook = ttk.Notebook(root)
        self.tab_calc = tk.Frame(self.notebook)
        self.tab_db = tk.Frame(self.notebook)
        self.notebook.add(self.tab_calc, text="Составление сметы")
        self.notebook.add(self.tab_db, text="Справочник")
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.setup_db_tab()
        self.setup_calc_tab()

    # -------------------------- Настройки --------------------------
    def open_settings(self):
        settings_win = tk.Toplevel(self.root)
        settings_win.title("Настройки")
        settings_win.geometry("600x250")
        settings_win.resizable(False, False)

        # Чтобы окно не уходило на задний план:
        settings_win.transient(self.root)  # Привязывает окно к родительскому (всегда сверху него)
        settings_win.grab_set()            # Делает окно модальным (блокирует клики по основному окну)

        db_folder_var = tk.StringVar(value=self.db_folder)
        export_folder_var = tk.StringVar(value=self.export_folder)

        frame_db = tk.LabelFrame(settings_win, text="Папка для базы данных", padx=10, pady=10)
        frame_db.pack(fill=tk.X, padx=10, pady=5)
        tk.Entry(frame_db, textvariable=db_folder_var, width=60).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_db, text="Обзор...", command=lambda: self.browse_folder(db_folder_var)).pack(side=tk.LEFT)

        frame_export = tk.LabelFrame(settings_win, text="Папка для сохранения смет", padx=10, pady=10)
        frame_export.pack(fill=tk.X, padx=10, pady=5)
        tk.Entry(frame_export, textvariable=export_folder_var, width=60).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_export, text="Обзор...", command=lambda: self.browse_folder(export_folder_var)).pack(side=tk.LEFT)

        btn_frame = tk.Frame(settings_win)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="Сохранить", bg="#4CAF50", fg="white", width=15,
                  command=lambda: self.save_settings_and_close(settings_win, db_folder_var.get(), export_folder_var.get())).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Отмена", width=15, command=settings_win.destroy).pack(side=tk.LEFT, padx=10)

    def browse_folder(self, var):
        folder = filedialog.askdirectory(title="Выберите папку")
        if folder:
            var.set(folder)

    def save_settings_and_close(self, win, db_folder, export_folder):
        os.makedirs(db_folder, exist_ok=True)
        os.makedirs(export_folder, exist_ok=True)

        self.db_folder = db_folder
        self.export_folder = export_folder
        self.db_file = os.path.join(self.db_folder, self.active_db_filename)

        save_settings({
            'db_folder': db_folder,
            'export_folder': export_folder,
            'active_db_filename': self.active_db_filename
        })

        self.db = self._load_db()
        self.refresh_db_table()
        self.update_combobox()
        self.refresh_db_list()  # Обновляем список после смены папки
        messagebox.showinfo("Готово", "Настройки сохранены.")
        win.destroy()

    # -------------------------- Работа с БД --------------------------
    def _load_db(self):
        if os.path.exists(self.db_file):
            try:
                temp_db = pd.read_excel(self.db_file)
                for c in COLS:
                    if c not in temp_db.columns:
                        temp_db[c] = "-"
                return temp_db[COLS]
            except Exception as e:
                messagebox.showerror("Ошибка БД", f"Не удалось загрузить базу:\n{e}")
                return pd.DataFrame(columns=COLS)
        return pd.DataFrame(columns=COLS)

    def setup_db_tab(self):
        # 🆕 Панель управления базами данных
        db_sel_frame = tk.Frame(self.tab_db)
        db_sel_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(db_sel_frame, text="📂 База данных:").pack(side=tk.LEFT)
        self.db_combo = ttk.Combobox(db_sel_frame, width=40, state="readonly")
        self.db_combo.pack(side=tk.LEFT, padx=5)
        tk.Button(db_sel_frame, text="➕ Создать", command=self.create_new_db).pack(side=tk.LEFT, padx=2)
        tk.Button(db_sel_frame, text="🗑 Удалить", command=self.delete_current_db).pack(side=tk.LEFT, padx=2)
        tk.Button(db_sel_frame, text="🔄 Обновить", command=self.refresh_db_list).pack(side=tk.LEFT, padx=2)
        self.db_combo.bind("<<ComboboxSelected>>", self.on_db_selected)
        self.refresh_db_list()

        frame_input = tk.LabelFrame(self.tab_db, text="Редактор базы", padx=10, pady=10)
        frame_input.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_input, text="Наименование работы:").grid(row=0, column=0, sticky="nw")
        self.work_text = tk.Text(frame_input, width=105, height=3, font=("Arial", 10))
        add_clipboard_support(self.work_text)
        add_context_menu(self.work_text)
        self.work_text.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="w")

        self.entries = {}
        fields = [
            ("Ед. изм. работы:", "Ед_изм_раб"),
            ("Материал:", "Материал"),
            ("Ед. изм. материала:", "Ед_изм"),
            ("Расход материала:", "Расход"),
            ("Цена материала:", "Цена_мат"),
            ("Цена работы (ед):", "Цена_раб")
        ]

        for i, (label, key) in enumerate(fields):
            r, c = divmod(i, 2)
            tk.Label(frame_input, text=label).grid(row=r + 1, column=c * 2, sticky="w")
            en = tk.Entry(frame_input, width=50)
            add_clipboard_support(en)
            add_context_menu(en)
            en.grid(row=r + 1, column=c * 2 + 1, padx=5, pady=2)
            self.entries[key] = en

        btn_f = tk.Frame(self.tab_db)
        btn_f.pack(pady=5)
        tk.Button(btn_f, text="Сохранить в базу", bg="#4CAF50", fg="white",
                  command=self.save_to_db).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_f, text="Удалить из базы", bg="#f44336", fg="white",
                  command=self.delete_from_db).pack(side=tk.LEFT, padx=5)

        # --- ФИЛЬТР ---
        filter_frame = tk.Frame(self.tab_db)
        filter_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(filter_frame, text="🔍 Поиск работы:").pack(side=tk.LEFT)
        self.filter_work = tk.Entry(filter_frame, width=35)
        self.filter_work.pack(side=tk.LEFT, padx=5)
        tk.Label(filter_frame, text="🔍 Поиск материала:").pack(side=tk.LEFT, padx=(10, 0))
        self.filter_mat = tk.Entry(filter_frame, width=35)
        self.filter_mat.pack(side=tk.LEFT, padx=5)
        tk.Button(filter_frame, text="Сброс", command=self.clear_db_filters).pack(side=tk.RIGHT, padx=5)
        
        # Живая фильтрация при вводе
        self.filter_work.bind('<KeyRelease>', lambda e: self.refresh_db_table())
        self.filter_mat.bind('<KeyRelease>', lambda e: self.refresh_db_table())

        # Таблица с НОВЫМ порядком колонок
        self.tree_db = ttk.Treeview(self.tab_db, columns=self.display_cols, show='headings')
        for c in self.display_cols:
            self.tree_db.heading(c, text=c, command=lambda _col=c: self.sort_column(_col))
            width = 250 if c == "Работа" else 100 if "Цена" in c else 80
            self.tree_db.column(c, width=width)
        self.tree_db.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree_db.bind("<<TreeviewSelect>>", self.load_to_entries)
        self.refresh_db_table()

    def clear_db_filters(self):
        self.filter_work.delete(0, tk.END)
        self.filter_mat.delete(0, tk.END)
        self.refresh_db_table()

    def sort_column(self, col):
        self.sort_orders[col] = not self.sort_orders[col]
        self.db = self.db.sort_values(by=col, ascending=self.sort_orders[col])
        self.refresh_db_table()

    def save_to_db(self):
        try:
            data = {c: self.entries[c].get().strip() for c in COLS if c != 'Работа'}
            data['Работа'] = self.work_text.get("1.0", tk.END).strip()

            if data['Материал'] in ("", "-", "0"):
                data['Материал'] = "-"

            for n_col in ['Расход', 'Цена_мат', 'Цена_раб']:
                data[n_col] = float(str(data[n_col]).replace(',', '.'))

            work_name = data['Работа']

            # Удаляем старую запись, если есть
            mask = (self.db['Работа'].str.strip() == work_name) & (self.db['Материал'].str.strip() == data['Материал'])
            self.db = self.db[~mask].reset_index(drop=True)
            self.db = pd.concat([self.db, pd.DataFrame([data])], ignore_index=True)

            os.makedirs(self.db_folder, exist_ok=True)
            self.db.to_excel(self.db_file, index=False)

            self.refresh_db_table()
            self.update_combobox()
            messagebox.showinfo("Готово", "Запись сохранена в базе.")
            self.refresh_estimate_if_needed(work_name)

        except ValueError:
            messagebox.showerror("Ошибка", "Проверьте числовые поля (Расход, Цена материала, Цена работы).")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить:\n{e}")

    def refresh_estimate_if_needed(self, work_name):
        # Проверка, есть ли эта работа в текущей смете
        works_in_smeta = {}
        for item in self.tree_smeta.get_children():
            vals = self.tree_smeta.item(item, 'values')
            if vals and str(vals[1]).startswith("Работа: "):
                name = vals[1].replace("Работа: ", "").strip()
                try:
                    vol = float(vals[4])
                except (ValueError, TypeError):
                    vol = 1.0
                works_in_smeta[name] = vol

        if work_name in works_in_smeta:
            self.tree_smeta.delete(*self.tree_smeta.get_children())
            for name, vol in works_in_smeta.items():
                self._add_work_to_smeta(name, vol, suppress_total_update=True)
            self.update_total_sum()

    def delete_from_db(self):
        sel = self.tree_db.selection()
        if not sel:
            messagebox.showwarning("Внимание", "Выберите строку для удаления.")
            return
        if not messagebox.askyesno("Подтверждение", "Удалить выбранную запись из базы?"):
            return

        try:
            idx_to_delete = int(sel[0])
            if idx_to_delete in self.db.index:
                self.db = self.db.drop(index=idx_to_delete).reset_index(drop=True)
                os.makedirs(self.db_folder, exist_ok=True)
                self.db.to_excel(self.db_file, index=False)
                self.refresh_db_table()
                self.update_combobox()
                messagebox.showinfo("Готово", "Запись удалена.")
            else:
                messagebox.showerror("Ошибка", "Индекс строки не найден в базе.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось удалить запись:\n{e}")

    def refresh_db_table(self):
        self.tree_db.delete(*self.tree_db.get_children())
        df = self.db.copy()

        # Применяем фильтры
        work_q = self.filter_work.get().strip().lower()
        mat_q = self.filter_mat.get().strip().lower()

        if work_q:
            mask = df['Работа'].astype(str).str.lower().str.contains(work_q, na=False)
            df = df[mask]
        if mat_q:
            mask = df['Материал'].astype(str).str.lower().str.contains(mat_q, na=False)
            df = df[mask]

        # Заполняем таблицу в порядке display_cols
        for idx, r in df.iterrows():
            vals = [r[c] for c in self.display_cols]
            self.tree_db.insert("", tk.END, iid=str(idx), values=vals)

    def load_to_entries(self, event):
        sel = self.tree_db.selection()
        if not sel:
            return
        try:
            # iid хранит оригинальный индекс DataFrame, что гарантирует точную привязку
            idx = int(sel[0])
            row = self.db.loc[idx]
        except (KeyError, ValueError):
            return

        self.work_text.delete("1.0", tk.END)
        self.work_text.insert("1.0", row['Работа'])
        for c in COLS[1:]:
            self.entries[c].delete(0, tk.END)
            self.entries[c].insert(0, row[c])

    # -------------------------- Вкладка "Смета" --------------------------
    def setup_calc_tab(self):
        # 👆 ВЕРХНЯЯ ПАНЕЛЬ: Выбор работы + Объем + Добавить
        frame_top = tk.Frame(self.tab_calc, pady=5)
        frame_top.pack(fill=tk.X, padx=10)

        self.work_combo = ttk.Combobox(frame_top, width=150) # Ширину НЕ меняем
        self.work_combo.pack(side=tk.LEFT, padx=5)
        self.work_combo.bind('<KeyRelease>', self.filter_works_combo)
        self.update_combobox()

        tk.Label(frame_top, text="Объем:").pack(side=tk.LEFT)
        self.vol_entry = tk.Entry(frame_top, width=10)
        add_clipboard_support(self.vol_entry)
        add_context_menu(self.vol_entry)
        self.vol_entry.pack(side=tk.LEFT, padx=5)

        tk.Button(frame_top, text="Добавить работу", bg="#2196F3", fg="white",
                  command=self.add_to_estimate).pack(side=tk.LEFT, padx=10)

        # 🛠 ПАНЕЛЬ ИНСТРУМЕНТОВ (ниже верхней)
        frame_tools = tk.Frame(self.tab_calc, pady=5)
        frame_tools.pack(fill=tk.X, padx=10)

        tk.Button(frame_tools, text="➕ Раздел", bg="#9C27B0", fg="white",
                  command=self.add_section).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_tools, text="➕ Материал", bg="#4CAF50", fg="white",
                  command=self.add_material_to_selected).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_tools, text="📋 Дубль мат.", bg="#FF9800", fg="white",
                  command=self.duplicate_material).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_tools, text="🗑 Удалить", bg="#f44336", fg="white",
                  command=self.remove_smeta_row).pack(side=tk.LEFT, padx=5)

        # 📊 ТАБЛИЦА СМЕТЫ
        self.calc_cols = ("№", "Наименование", "Ед. изм.", "Норма", "Кол-во", "Цена", "Стоимость")
        self.tree_smeta = ttk.Treeview(self.tab_calc, columns=self.calc_cols, show='headings', selectmode='extended')
        for i, c in enumerate(self.calc_cols):
            self.tree_smeta.heading(c, text=c)
            self.tree_smeta.column(c, width=100)
        self.tree_smeta.column("Наименование", width=450)
        self.tree_smeta.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree_smeta.tag_configure("section", background="#E1BEE7", font=("Arial", 10, "bold"))
        self.tree_smeta.bind("<Double-1>", self.on_tree_double_click)
        self.tree_smeta.bind("<Control-z>", self.undo_action)
        self.tree_smeta.bind("<Button-3>", self._show_context_menu)

        # 👇 НИЖНЯЯ ПАНЕЛЬ: Загрузить + Выгрузить
        frame_bottom = tk.Frame(self.tab_calc, pady=5)
        frame_bottom.pack(fill=tk.X, padx=10)

        tk.Button(frame_bottom, text="📂 Загрузить смету", bg="#607D8B", fg="white",
                  command=self.load_estimate).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_bottom, text="Выгрузить в Excel", bg="#FF9800", fg="white",
                  command=self.export_excel).pack(side=tk.LEFT, padx=5)

        # 📉 ИТОГ (в самом низу)
        total_frame = tk.Frame(self.tab_calc)
        total_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        tk.Label(total_frame, text=" ").pack(side=tk.LEFT, expand=True)
        self.total_label = tk.Label(
            total_frame,
            text="ИТОГО ПО СМЕТЕ: 0.00 ₽",
            font=("Arial", 14, "bold"),
            fg="#1565C0"
        )
        self.total_label.pack(side=tk.RIGHT)

    def add_section(self):
        """Добавляет строку-раздел в конец сметы."""
        self.tree_smeta.insert("", tk.END, values=("", "РАЗДЕЛ: Новый раздел", "", "", "", "", ""), tags=("section",))

    def on_tree_double_click(self, event):
        """Обработка двойного клика: разделы, работы, материалы."""
        item = self.tree_smeta.identify_row(event.y)
        if not item: return
        
        col = self.tree_smeta.identify_column(event.x)
        if not col or col == '#0': return
        
        col_idx = int(col[1:]) - 1
        vals = self.tree_smeta.item(item, 'values')
        if not vals: return
        
        name_raw = str(vals[1]).strip()
        
        # 1. Редактирование названия раздела
        if name_raw.startswith("РАЗДЕЛ: ") and col_idx == 1:
            self._edit_section_inline(item, vals)
            return
            
        # 2. Редактирование ячеек сметы (1-Наименование, 3-Норма/Расход, 5-Цена)
        if "ИТОГО" not in name_raw and col_idx in [1, 3, 5]:
            self._start_cell_edit(item, vals, col_idx)

    def update_combobox(self):
        if not self.db.empty:
            self.all_works_list = sorted(list(self.db['Работа'].unique()))
            self.work_combo['values'] = self.all_works_list
        else:
            self.all_works_list = []
            self.work_combo['values'] = []

    def filter_works_combo(self, event=None):
        """Автоматическая фильтрация списка работ при вводе в ComboBox."""
        typed = self.work_combo.get().strip()
        if not hasattr(self, 'all_works_list'): return

        if not typed:
            self.work_combo['values'] = self.all_works_list
            return

        # Поиск без учёта регистра
        lower_typed = typed.lower()
        filtered = [w for w in self.all_works_list if lower_typed in w.lower()]
        self.work_combo['values'] = filtered
        self.work_combo.icursor(tk.END)  # Удерживаем курсор в конце строки

    def refresh_db_list(self):
        """Обновляет список доступных баз в выпадающем меню."""
        if not os.path.exists(self.db_folder): os.makedirs(self.db_folder, exist_ok=True)
        files = sorted([f for f in os.listdir(self.db_folder) if f.lower().endswith('.xlsx')])
        self.db_combo['values'] = files
        if self.active_db_filename in files:
            self.db_combo.set(self.active_db_filename)
        elif files:
            self.db_combo.set(files[0])
            self.on_db_selected()

    def on_db_selected(self, event=None):
        """Загружает выбранную базу данных."""
        new_db = self.db_combo.get()
        if new_db and new_db != self.active_db_filename:
            self.active_db_filename = new_db
            self.db_file = os.path.join(self.db_folder, new_db)
            self._save_active_db_to_settings()
            self.db = self._load_db()
            self.refresh_db_table()
            self.update_combobox()

    def _save_active_db_to_settings(self):
        """Сохраняет имя активной базы в settings.json."""
        settings = load_settings()
        settings['active_db_filename'] = self.active_db_filename
        save_settings(settings)

    def create_new_db(self):
        """Создает новую пустую базу и переключается на неё."""
        new_name = simpledialog.askstring("Новая база", "Введите имя файла (например, База_Кровля.xlsx):")
        if not new_name: return
        if not new_name.lower().endswith('.xlsx'): new_name += '.xlsx'
        new_path = os.path.join(self.db_folder, new_name)
        if os.path.exists(new_path):
            return messagebox.showwarning("Внимание", "Файл уже существует.")
        pd.DataFrame(columns=COLS).to_excel(new_path, index=False)
        self.refresh_db_list()
        self.db_combo.set(new_name)
        self.on_db_selected()
        messagebox.showinfo("Готово", f"База '{new_name}' создана и загружена.")

    def delete_current_db(self):
        """Удаляет текущую базу данных."""
        if not self.active_db_filename: return
        if not messagebox.askyesno("Подтверждение", f"Удалить базу '{self.active_db_filename}'?\nЭто действие нельзя отменить."): return
        path = os.path.join(self.db_folder, self.active_db_filename)
        try:
            os.remove(path)
            self.refresh_db_list()
            if self.db_combo['values']:
                self.db_combo.set(self.db_combo['values'][0])
                self.on_db_selected()
            else:
                self.db = pd.DataFrame(columns=COLS)
                self.refresh_db_table()
                self.update_combobox()
            messagebox.showinfo("Готово", "База удалена.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось удалить файл:\n{e}")

    def add_to_estimate(self):
        work_name = self.work_combo.get().strip()
        if not work_name:
            messagebox.showwarning("Внимание", "Выберите работу из списка.")
            return
        try:
            vol = float(self.vol_entry.get().replace(',', '.'))
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректное число в поле 'Объем'.")
            return

        self._add_work_to_smeta(work_name, vol)
        self.update_total_sum()

    def _add_work_to_smeta(self, work_name, volume, suppress_total_update=False):
        items = self.db[self.db['Работа'].str.strip() == work_name]
        if items.empty:
            messagebox.showwarning("Внимание", f"Работа '{work_name}' не найдена в базе!")
            return

        unit_w = items.iloc[0]['Ед_изм_раб']
        price_w = float(items.iloc[0]['Цена_раб'])

        next_num = 1
        for item in self.tree_smeta.get_children():
            vals = self.tree_smeta.item(item, 'values')
            if vals and vals[0] != "" and str(vals[0]).isdigit():
                next_num = max(next_num, int(vals[0]) + 1)

        work_total = round(volume * price_w, 2)
        self.tree_smeta.insert("", tk.END, values=(
            next_num, f"Работа: {work_name}", unit_w, "-", volume, price_w, work_total
        ))

        mat_total = 0.0
        for _, row in items.iterrows():
            mat_name = str(row['Материал']).strip()
            if mat_name not in ("", "-", "0"):
                qty = row['Расход'] * volume
                m_cost = round(qty * row['Цена_мат'], 2)
                mat_total += m_cost
                self.tree_smeta.insert("", tk.END, values=(
                    "", f"   > {mat_name}", row['Ед_изм'],
                    row['Расход'], round(qty, 3), row['Цена_мат'], m_cost
                ))

        self.tree_smeta.insert("", tk.END, values=(
            "", f"ИТОГО ПО УЗЛУ: {work_name}", "", "", "", "Сумма:",
            round(work_total + mat_total, 2)
        ))

        if not suppress_total_update:
            self.update_total_sum()

    def remove_smeta_row(self):
        selected = self.tree_smeta.selection()
        if not selected:
            return

        for item in selected:
            vals = self.tree_smeta.item(item, 'values')
            if vals and "ИТОГО ПО УЗЛУ" in str(vals):
                messagebox.showwarning("Внимание", "Нельзя удалить итоговую строку напрямую.\n"
                                                   "Для удаления узла целиком удалите строку с работой.")
                return

        for item in selected:
            self.tree_smeta.delete(item)

        self.recalculate_smeta()

    def is_material_row(self, name: str) -> bool:
        return name.strip().startswith(">")

    def recalculate_smeta(self):
        # 1. Собираем строки вместе с их тегами (разделами)
        rows = []
        for item in self.tree_smeta.get_children():
            vals = self.tree_smeta.item(item, 'values')
            tags = self.tree_smeta.item(item, 'tags') or ()  # Гарантируем, что tags всегда кортеж
            if vals and "ИТОГО ПО УЗЛУ" not in str(vals):
                rows.append((vals, tags))

        self.tree_smeta.delete(*self.tree_smeta.get_children())

        node_work_name = None
        node_work_cost = 0.0
        node_mat_cost = 0.0
        next_num = 1

        # 2. Восстанавливаем дерево с сохранением тегов
        for vals, tags in rows:
            name = str(vals[1])

            # Разделы просто вставляем обратно без изменений
            if name.startswith("РАЗДЕЛ: "):
                self.tree_smeta.insert("", tk.END, values=vals, tags=tags)
                continue

            if name.startswith("Работа: "):
                # Завершаем предыдущий узел, если он был
                if node_work_name is not None:
                    total = node_work_cost + node_mat_cost
                    self.tree_smeta.insert("", tk.END, values=(
                        "", f"ИТОГО ПО УЗЛУ: {node_work_name}", "", "", "", "Сумма:", round(total, 2)
                    ))
                node_work_name = name.replace("Работа: ", "")
                node_work_cost = float(vals[6]) if vals[6] else 0.0
                node_mat_cost = 0.0
                new_vals = list(vals)
                new_vals[0] = next_num
                self.tree_smeta.insert("", tk.END, values=tuple(new_vals))
                next_num += 1
            else:
                new_vals = list(vals)
                new_vals[0] = ""
                self.tree_smeta.insert("", tk.END, values=tuple(new_vals))
                cost = float(vals[6]) if vals[6] else 0.0
                node_mat_cost += cost

        # Закрываем последний узел
        if node_work_name is not None:
            total = node_work_cost + node_mat_cost
            self.tree_smeta.insert("", tk.END, values=(
                "", f"ИТОГО ПО УЗЛУ: {node_work_name}", "", "", "", "Сумма:", round(total, 2)
            ))

        self.update_total_sum()

    def update_total_sum(self):
        total = 0.0
        for item in self.tree_smeta.get_children():
            vals = self.tree_smeta.item(item, 'values')
            if vals and "ИТОГО ПО УЗЛУ" in str(vals):
                try:
                    total += float(vals[6])
                except (ValueError, TypeError):
                    pass
        # Форматируем число с разделителем тысяч пробелом
        formatted_total = format(total, ',.2f').replace(',', ' ')
        self.total_label.config(text=f"ИТОГО ПО СМЕТЕ: {formatted_total} ₽")
        self.root.update_idletasks()

    def _edit_section_inline(self, item, vals):
        x, y, w, h = self.tree_smeta.bbox(item, column='#2')
        if not (x, y, w, h): return
        root_x = self.tree_smeta.winfo_rootx() - self.root.winfo_rootx()
        root_y = self.tree_smeta.winfo_rooty() - self.root.winfo_rooty()
        self.edit_entry = tk.Entry(self.root, font=("Arial", 10), bd=1, relief="solid")
        self.edit_entry.place(x=root_x + x, y=root_y + y, width=w, height=h)

        # 🔧 Явно подключаем поддержку Ctrl+C/V и ПКМ для динамического поля
        add_clipboard_support(self.edit_entry)
        add_context_menu(self.edit_entry)

        self.edit_entry.insert(0, vals[1].replace("РАЗДЕЛ: ", ""))
        self.edit_entry.focus(); self.edit_entry.select_range(0, tk.END)
        self.edit_item, self.edit_orig, self.edit_col_idx = item, list(vals), 1
        self.edit_entry.bind("<Return>", lambda e: self._save_section_edit())
        self.edit_entry.bind("<Escape>", lambda e: self._destroy_edit())
        self.edit_entry.bind("<FocusOut>", lambda e: self.root.after(100, self._save_section_edit))

    def _save_section_edit(self):
        if self.edit_entry is None or not self.edit_entry.winfo_exists(): return
        nm = self.edit_entry.get().strip()
        if nm:
            nv = self.edit_orig[:]; nv[1] = f"РАЗДЕЛ: {nm}"
            self.tree_smeta.item(self.edit_item, values=tuple(nv))
        self._destroy_edit()

    def _start_cell_edit(self, item, vals, col_idx):
        # Сохраняем состояние для отмены (Ctrl+Z)
        self.undo_stack.append((item, col_idx, vals[col_idx]))
        
        ct = f'#{col_idx+1}'
        bbox = self.tree_smeta.bbox(item, column=ct)
        if not bbox: return
        x, y, w, h = bbox
        root_x = self.tree_smeta.winfo_rootx() - self.root.winfo_rootx()
        root_y = self.tree_smeta.winfo_rooty() - self.root.winfo_rooty()
        self.edit_entry = tk.Entry(self.root, font=("Arial", 10), bd=1, relief="solid")
        self.edit_entry.place(x=root_x + x, y=root_y + y, width=w, height=h)

        # 🔧 Явно подключаем поддержку Ctrl+C/V и ПКМ
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
        if self.edit_entry is None or not self.edit_entry.winfo_exists(): return
        nv = self.edit_entry.get().strip()
        if not nv: return self._destroy_edit()

        nvals = self.edit_orig[:]
        if self.edit_col_idx in [3, 5]:
            try: 
                nvals[self.edit_col_idx] = float(nv.replace(',', '.'))
            except: 
                messagebox.showerror("Ошибка", "Введите корректное число.")
                self.edit_entry.focus()
                return
        else:
            nvals[self.edit_col_idx] = nv

        # 🔧 Фикс 2: Восстанавливаем префикс отступа при редактировании названия
        if self.edit_col_idx == 1:
            orig = str(self.edit_orig[1]).strip()
            clean = nvals[1].replace("Работа:", "").replace(">", "").strip()
            if orig.startswith(">"): nvals[1] = f"   > {clean}"
            elif orig.startswith("Работа:"): nvals[1] = f"Работа: {clean}"

        self.tree_smeta.item(self.edit_item, values=tuple(nvals))
        self._sync_db(nvals, self.edit_col_idx)

        # 🔧 Фикс 3: Если изменилась цена материала -> обновляем все такие же материалы в смете
        if self.edit_col_idx == 5 and str(self.edit_orig[1]).strip().startswith(">"):
            mat_name = str(self.edit_orig[1]).replace(">", "").replace("   ", "").strip()
            new_price = nvals[5]
            for item in self.tree_smeta.get_children():
                if item == self.edit_item: continue
                vals = list(self.tree_smeta.item(item, 'values'))
                v_name = str(vals[1]).strip()
                if v_name.startswith(">") and mat_name == v_name.replace(">", "").replace("   ", "").strip():
                    vals[5] = new_price
                    qty = float(vals[4]) if vals[4] else 0.0
                    vals[6] = round(qty * float(new_price), 2)
                    self.tree_smeta.item(item, values=tuple(vals))

        # Если изменилась цена работы -> обновляем все такие же работы в смете
        if self.edit_col_idx == 5 and str(self.edit_orig[1]).strip().startswith("Работа:"):
            work_name = str(self.edit_orig[1]).replace("Работа:", "").strip()
            new_price = nvals[5]
            for item in self.tree_smeta.get_children():
                if item == self.edit_item: continue
                vals = list(self.tree_smeta.item(item, 'values'))
                v_name = str(vals[1]).strip()
                if v_name.startswith("Работа:") and work_name == v_name.replace("Работа:", "").strip():
                    vals[5] = new_price
                    vol = float(vals[4]) if vals[4] else 0.0
                    vals[6] = round(vol * float(new_price), 2)
                    self.tree_smeta.item(item, values=tuple(vals))

        try:
            self._recalculate_all_costs()
        except Exception as e:
            messagebox.showerror("Ошибка пересчёта", f"Не удалось обновить итоги:\n{e}")
        finally:
            self._destroy_edit()

    def _recalculate_all_costs(self):
        """Двухпроходный пересчёт: 1. Кол-во и стоимость строк, 2. Итоги по узлам."""
        work_vol = 1.0

        # --- ПРОХОД 1: Обновляем количество материалов и стоимость всех работ/материалов ---
        for item in self.tree_smeta.get_children():
            vals = list(self.tree_smeta.item(item, 'values'))
            name = str(vals[1]).strip()

            if not name or "РАЗДЕЛ" in name or "ИТОГО" in name:
                continue

            if name.startswith("Работа: "):
                # Запоминаем объем текущего узла для последующих материалов
                try: work_vol = float(vals[4]) if vals[4] else 1.0
                except: work_vol = 1.0
                
                # 🔧 FIX: Пересчитываем стоимость работы = Объем × Цена
                try:
                    vol = float(vals[4]) if vals[4] else 0.0
                    price_str = str(vals[5]).replace(',', '.').strip()
                    price = float(price_str) if price_str and price_str not in ("-", "Сумма:", "заказчика") else 0.0
                    vals[6] = round(vol * price, 2)
                except: vals[6] = 0.0
                self.tree_smeta.item(item, values=tuple(vals))

            elif self.is_material_row(name):
                # Пересчитываем количество материала на основе нормы и объема работы
                try:
                    norm = float(vals[3]) if vals[3] else 0.0
                    vals[4] = round(norm * work_vol, 3)
                except: pass
                
                # Пересчитываем стоимость материала = Кол-во × Цена
                try:
                    qty = float(vals[4]) if vals[4] else 0.0
                    price_str = str(vals[5]).replace(',', '.').strip()
                    price = float(price_str) if price_str and price_str not in ("-", "Сумма:", "заказчика") else 0.0
                    vals[6] = round(qty * price, 2)
                except: vals[6] = 0.0
                self.tree_smeta.item(item, values=tuple(vals))

        # --- ПРОХОД 2: Пересчитываем строки "ИТОГО ПО УЗЛУ" ---
        node_work_cost = 0.0
        node_mat_cost = 0.0

        for item in self.tree_smeta.get_children():
            vals = list(self.tree_smeta.item(item, 'values'))
            name = str(vals[1]).strip()

            if not name or "РАЗДЕЛ" in name: continue

            if name.startswith("Работа: "):
                node_work_cost = float(vals[6]) if vals[6] else 0.0
                node_mat_cost = 0.0
            elif self.is_material_row(name):
                node_mat_cost += float(vals[6]) if vals[6] else 0.0
            elif "ИТОГО ПО УЗЛУ" in name:
                total = round(node_work_cost + node_mat_cost, 2)
                vals[6] = total
                self.tree_smeta.item(item, values=tuple(vals))
                # Сброс после подытога (для страховки)
                node_work_cost = 0.0
                node_mat_cost = 0.0

        # 🔧 Фикс 4: Обновляем нижний виджет ИТОГО
        self.update_total_sum()

    def _sync_db(self, nv, col_idx):
        ov = str(self.edit_orig[1]).strip()
        is_work = "Работа:" in ov
        is_mat = ">" in ov
        clean = ov.replace("Работа:", "").replace(">", "").strip()
        try:
            if col_idx == 1:
                nc = nv[1].replace("Работа:", "").replace(">", "").strip()
                if is_work: self.db.loc[self.db['Работа'].str.strip()==clean, 'Работа'] = nc
                elif is_mat: self.db.loc[self.db['Материал'].str.strip()==clean, 'Материал'] = nc
            elif col_idx == 3 and is_mat:
                self.db.loc[self.db['Материал'].str.strip()==clean, 'Расход'] = float(nv[3])
            elif col_idx == 5:
                p = float(nv[5])
                if is_work: self.db.loc[self.db['Работа'].str.strip()==clean, 'Цена_раб'] = p
                elif is_mat: self.db.loc[self.db['Материал'].str.strip()==clean, 'Цена_мат'] = p
            self.db.to_excel(self.db_file, index=False)
            self.refresh_db_table()
            self.update_combobox()
        except Exception as e: messagebox.showerror("Ошибка БД", str(e))

    def _recalc_costs(self):
        for item in self.tree_smeta.get_children():
            vals = list(self.tree_smeta.item(item, 'values'))
            if not vals or "ИТОГО" in str(vals[1]) or "РАЗДЕЛ" in str(vals[1]): continue
            try: vals[6] = round((float(vals[4] or 0)) * (float(vals[5] or 0)), 2)
            except: pass
            self.tree_smeta.item(item, values=tuple(vals))
        self.update_total_sum()

    def _destroy_edit(self):
        if self.edit_entry is not None and self.edit_entry.winfo_exists(): self.edit_entry.destroy()
        self.edit_entry = None
        self.edit_item = None
        self.edit_col_idx = None
        self.edit_orig = None

    def _open_material_dialog(self, initial_data=None):
        """Модальное окно для ввода данных материала."""
        win = tk.Toplevel(self.root)
        win.title("Добавить/Изменить материал")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        win.geometry("350x180")

        result = {}
        entries = {}
        fields = [("Наименование:", "name", 25), ("Ед. изм.:", "unit", 8),
                  ("Норма расхода:", "norm", 10), ("Цена за ед.:", "price", 10)]

        for i, (lbl, key, w) in enumerate(fields):
            tk.Label(win, text=lbl).grid(row=i, column=0, padx=5, pady=5, sticky="e")
            e = tk.Entry(win, width=w)
            e.grid(row=i, column=1, padx=5, pady=5, sticky="w")
            if initial_data and key in initial_data:
                e.insert(0, str(initial_data[key]))
            entries[key] = e

        def apply():
            for k, e in entries.items(): result[k] = e.get().strip()
            win.destroy()
        def cancel(): result.clear(); win.destroy()

        btn_f = tk.Frame(win)
        btn_f.grid(row=4, column=0, columnspan=2, pady=10)
        tk.Button(btn_f, text="Добавить", bg="#4CAF50", fg="white", command=apply).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_f, text="Отмена", command=cancel).pack(side=tk.LEFT, padx=5)

        win.wait_window()
        return result if result else None

    def add_material_to_selected(self):
        """Добавляет строку материала к выбранной работе или после материала."""
        sel = self.tree_smeta.selection()
        if not sel:
            return messagebox.showwarning("Внимание", "Выберите работу или материал.")
        item = sel[0]
        vals = self.tree_smeta.item(item, 'values')
        name_raw = str(vals[1]).strip()

        if name_raw.startswith("РАЗДЕЛ:"):
            return messagebox.showwarning("Внимание", "Нельзя добавить материал в раздел. Выберите работу.")

        children = list(self.tree_smeta.get_children())
        insert_idx = children.index(item) + 1

        data = self._open_material_dialog()
        if not data or not data.get('name'): return

        # Вставляем строку: №="", Имя="> ...", Ед, Норма, Кол-во="", Цена, Стоимость=""
        self.tree_smeta.insert("", insert_idx, values=("", f"   > {data['name']}", data['unit'], data['norm'], "", data['price'], ""))
        self._recalculate_all_costs() # ✅ Мгновенный пересчёт Кол-во, Стоимости и Итогов

    def duplicate_material(self):
        """Дублирует выбранную строку материала с сохранением всех данных."""
        sel = self.tree_smeta.selection()
        if not sel:
            return messagebox.showwarning("Внимание", "Выберите строку материала.")
        item = sel[0]
        vals = list(self.tree_smeta.item(item, 'values'))
        name_raw = str(vals[1]).strip()
        if not name_raw.startswith(">"):
            return messagebox.showwarning("Внимание", "Выбрана не строка материала.")

        children = list(self.tree_smeta.get_children())
        insert_idx = children.index(item) + 1
        self.tree_smeta.insert("", insert_idx, values=tuple(vals))
        self._recalculate_all_costs() # ✅ Пересчёт сработает автоматически

        # -------------------------- Импорт / Экспорт --------------------------
    def load_estimate(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл сметы",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            xl = pd.ExcelFile(file_path)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")
            return

        meta_df = None
        if "Meta" in xl.sheet_names:
            meta_df = pd.read_excel(xl, sheet_name="Meta")
        else:
            if not messagebox.askyesno("Предупреждение",
                                       "В файле отсутствует служебный лист Meta.\n"
                                       "Восстановление сметы может быть неполным.\n"
                                       "Продолжить?"):
                return

        if "Смета" not in xl.sheet_names:
            messagebox.showerror("Ошибка", "В файле отсутствует лист 'Смета'.")
            return

        # Читаем без заголовка, чтобы индексы точно соответствовали колонкам Excel
        smeta_df = pd.read_excel(xl, sheet_name="Смета", header=None)

        # --- Обработка Meta (обновление справочника) ---
        if meta_df is not None and not meta_df.empty:
            new_rows_list = []
            required_cols = ['Работа', 'Материал', 'Ед_изм_раб', 'Ед_изм', 'Расход', 'Цена_мат', 'Цена_раб']
            for col in required_cols:
                if col not in meta_df.columns:
                    meta_df[col] = "-"

            for _, meta_row in meta_df.iterrows():
                work = str(meta_row['Работа']).strip()
                material = str(meta_row['Материал']).strip()

                existing = self.db[
                    (self.db['Работа'].astype(str).str.strip() == work) &
                    (self.db['Материал'].astype(str).str.strip() == material)
                ]
                if existing.empty:
                    new_rows_list.append(meta_row)
                else:
                    match_found = False
                    for _, exist_row in existing.iterrows():
                        try:
                            if (str(exist_row['Ед_изм_раб']).strip() == str(meta_row['Ед_изм_раб']).strip() and
                                str(exist_row['Ед_изм']).strip() == str(meta_row['Ед_изм']).strip() and
                                abs(float(exist_row['Расход']) - float(meta_row['Расход'])) < 1e-6 and
                                abs(float(exist_row['Цена_мат']) - float(meta_row['Цена_мат'])) < 1e-6 and
                                abs(float(exist_row['Цена_раб']) - float(meta_row['Цена_раб'])) < 1e-6):
                                match_found = True
                                break
                        except (ValueError, TypeError):
                            continue
                    if not match_found:
                        new_rows_list.append(meta_row)

            if new_rows_list:
                if messagebox.askyesno("Новые данные",
                                       f"В смете найдено {len(new_rows_list)} новых или изменённых записей.\n"
                                       "Добавить их в справочник?"):
                    new_rows_df = pd.DataFrame(new_rows_list)
                    for num_col in ['Расход', 'Цена_мат', 'Цена_раб']:
                        if num_col in new_rows_df.columns:
                            new_rows_df[num_col] = pd.to_numeric(new_rows_df[num_col], errors='coerce').fillna(0.0)
                    self.db = pd.concat([self.db, new_rows_df], ignore_index=True)
                    os.makedirs(self.db_folder, exist_ok=True)
                    self.db.to_excel(self.db_file, index=False)
                    self.refresh_db_table()
                    self.update_combobox()
                    messagebox.showinfo("Готово", "Новые записи добавлены в справочник.")
            else:
                messagebox.showinfo("Информация", "Все записи из сметы уже присутствуют в справочнике.")

        # --- Загрузка строк сметы в интерфейс ---
        self.tree_smeta.delete(*self.tree_smeta.get_children())

        # Пропускаем строку 0 (заголовки Excel), начинаем с 1
        for idx, row in smeta_df.iloc[1:].iterrows():
            row = smeta_df.iloc[idx]
            try:
                col0 = row[0]  # № п/п
                col1 = row[1]  # Наименование
                col4 = row[4]  # Кол-во
            except IndexError:
                continue

            if pd.isna(col1):
                continue

            name_raw = str(col1).strip()
            if not name_raw or "ИТОГО" in name_raw or "в том числе" in name_raw:
                continue

            # 👇 ЗАГРУЗКА РАЗДЕЛОВ
            if name_raw.startswith("РАЗДЕЛ: "):
                self.tree_smeta.insert("", tk.END, values=("", name_raw, "", "", "", "", ""), tags=("section",))
                continue

            # Работа определяется по наличию номера в первой колонке
            if pd.notna(col0) and str(col0).strip().replace('.', '').isdigit():
                work_name = name_raw.replace("Работа: ", "").strip()
                try:
                    vol = float(col4) if pd.notna(col4) else 0.0
                except ValueError:
                    vol = 0.0
                self._add_work_to_smeta(work_name, vol)

        self.update_total_sum()
        messagebox.showinfo("Готово", "Смета загружена.")

    def export_excel(self):
        try:
            now = datetime.now()
            timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
            export_path = os.path.join(self.export_folder, f"Smeta_{timestamp}.xlsx")

            writer = pd.ExcelWriter(export_path, engine='xlsxwriter')
            wb = writer.book
            ws = wb.add_worksheet("Смета")

            # --- Форматы ---
            header_fmt = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center', 'text_wrap': True, 'font_name': 'Arial', 'font_size': 12})
            work_first_fmt = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC', 'text_wrap': True, 'font_name': 'Arial', 'font_size': 10})
            work_first_num_fmt = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC', 'num_format': '#,##0.00', 'align': 'right', 'font_name': 'Arial', 'font_size': 10})
            mat_first_fmt = wb.add_format({'border': 1, 'bg_color': '#DDEBF7', 'indent': 1, 'text_wrap': True, 'font_name': 'Arial', 'font_size': 10})
            mat_first_num_fmt = wb.add_format({'border': 1, 'bg_color': '#DDEBF7', 'num_format': '#,##0.00', 'align': 'right', 'font_name': 'Arial', 'font_size': 10})
            work_std_fmt = wb.add_format({'bold': True, 'border': 1, 'text_wrap': True, 'font_name': 'Arial', 'font_size': 10})
            work_std_num_fmt = wb.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00', 'align': 'right', 'font_name': 'Arial', 'font_size': 10})
            mat_std_fmt = wb.add_format({'border': 1, 'indent': 1, 'text_wrap': True, 'font_name': 'Arial', 'font_size': 10})
            mat_std_num_fmt = wb.add_format({'border': 1, 'num_format': '#,##0.00', 'align': 'right', 'font_name': 'Arial', 'font_size': 10})
            subtotal_fmt = wb.add_format({'bold': True, 'border': 1, 'align': 'right', 'text_wrap': True, 'font_name': 'Arial', 'font_size': 10})
            final_fmt = wb.add_format({'bold': True, 'border': 1, 'bg_color': '#FFCC00', 'align': 'right', 'num_format': '#,##0.00', 'font_name': 'Arial', 'font_size': 12})
            total_label_fmt = wb.add_format({'bold': True, 'border': 1, 'align': 'left', 'text_wrap': True, 'font_name': 'Arial', 'font_size': 12})
            section_fmt = wb.add_format({'bold': True, 'bg_color': '#E1BEE7', 'align': 'left', 'valign': 'vcenter', 'font_size': 11, 'border': 1})

            headers = ["№ п/п", "Наименование работ и затрат", "Ед. изм.", "Нормы расхода", "Кол-во", "Цена за един. с НДС, руб.", "Стоимость с НДС, руб."]
            for col_idx, h in enumerate(headers): ws.write(0, col_idx, h, header_fmt)
            ws.set_column(0, 0, 8); ws.set_column(1, 1, 80); ws.set_column(2, 3, 14); ws.set_column(4, 6, 16)

            seen_works = set()
            seen_materials = set()
            first_work_price_row = {}  # 🔧 FIX 2: Хранит номер строки Excel с первой ценой работы
            first_mat_price_row = {}   # 🔧 FIX 2: Хранит номер строки Excel с первой ценой материала

            all_items = self.tree_smeta.get_children()
            rows = [self.tree_smeta.item(item, 'values') for item in all_items]

            excel_row = 1
            total_formula_parts = []
            i = 0

            while i < len(rows):
                vals = rows[i]
                name_raw = str(vals[1])

                if name_raw.startswith("РАЗДЕЛ:"):
                    section_title = name_raw.replace("РАЗДЕЛ:", "").strip()
                    ws.merge_range(excel_row, 0, excel_row, 6, section_title, section_fmt)
                    excel_row += 1; i += 1; continue

                if name_raw.startswith("Работа:"):
                    work_name = name_raw.replace("Работа:", "").strip()
                    unit_w = str(vals[2])
                    vol = float(vals[4]) if vals[4] else 0.0
                    price_w = float(vals[5]) if vals[5] else 0.0

                    # 🔧 FIX 1: Преобразуем № в int, чтобы Excel записал число, а не текст '1
                    try: work_num = int(vals[0])
                    except: work_num = 0

                    if work_name not in seen_works:
                        seen_works.add(work_name)
                        first_work_price_row[work_name] = excel_row + 1
                        w_fmt, w_num_fmt = work_first_fmt, work_first_num_fmt
                        ws.write(excel_row, 5, price_w, w_num_fmt) # Первая цена: статичное значение
                    else:
                        w_fmt, w_num_fmt = work_std_fmt, work_std_num_fmt
                        ref_row = first_work_price_row[work_name]
                        ws.write_formula(excel_row, 5, f"=F{ref_row}", w_num_fmt) # 🔧 FIX 2: Повторная цена: ссылка на первую

                    cost_formula = f"=E{excel_row+1}*F{excel_row+1}"
                    ws.write(excel_row, 0, work_num, w_fmt)
                    ws.write(excel_row, 1, work_name, w_fmt)
                    ws.write(excel_row, 2, unit_w, w_fmt)
                    ws.write(excel_row, 3, "-", w_fmt)
                    ws.write(excel_row, 4, vol, w_num_fmt)
                    ws.write_formula(excel_row, 6, cost_formula, w_num_fmt)

                    work_excel_row = excel_row
                    excel_row += 1; i += 1

                    while i < len(rows):
                        next_vals = rows[i]
                        next_name = str(next_vals[1])
                        if next_name.startswith(">") or next_name.startswith("   >") or next_name.startswith("    >"):
                            mat_name = next_name.replace(">", "").replace("   ", "").strip()
                            mat_price_val = float(next_vals[5]) if next_vals[5] else 0.0

                            if mat_name not in seen_materials:
                                seen_materials.add(mat_name)
                                first_mat_price_row[mat_name] = excel_row + 1
                                m_fmt, m_num_fmt = mat_first_fmt, mat_first_num_fmt
                                ws.write(excel_row, 5, mat_price_val, m_num_fmt) # Первая цена: статичное значение
                            else:
                                m_fmt, m_num_fmt = mat_std_fmt, mat_std_num_fmt
                                ref_row = first_mat_price_row[mat_name]
                                ws.write_formula(excel_row, 5, f"=F{ref_row}", m_num_fmt) # 🔧 FIX 2: Повторная цена: ссылка на первую

                            qty_formula = f"=E{work_excel_row+1}*D{excel_row+1}"
                            mat_cost_formula = f"=E{excel_row+1}*F{excel_row+1}"

                            ws.write(excel_row, 0, "", m_fmt)
                            ws.write(excel_row, 1, mat_name, m_fmt)
                            ws.write(excel_row, 2, str(next_vals[2]), m_fmt)
                            ws.write(excel_row, 3, float(next_vals[3]) if next_vals[3] else 0.0, m_num_fmt)
                            ws.write_formula(excel_row, 4, qty_formula, m_num_fmt)
                            # Цена уже записана выше
                            ws.write_formula(excel_row, 6, mat_cost_formula, m_num_fmt)

                            excel_row += 1; i += 1
                        else:
                            break

                    sum_range = f"G{work_excel_row+1}:G{excel_row}"
                    subtotal_formula = f"=SUM({sum_range})"
                    ws.merge_range(excel_row, 1, excel_row, 4, f"ИТОГО ПО УЗЛУ: {work_name}", subtotal_fmt)
                    ws.write(excel_row, 5, "Сумма:", subtotal_fmt)
                    ws.write_formula(excel_row, 6, subtotal_formula, subtotal_fmt)
                    total_formula_parts.append(f"G{excel_row+1}")
                    excel_row += 1
                else:
                    i += 1

            # --- Общие итоги ---
            row_total = excel_row + 2
            ws.write(row_total, 1, "ИТОГО, с учетом НДС 22%:", total_label_fmt)
            total_formula = "=" + "+".join(total_formula_parts) if total_formula_parts else "0"
            ws.write_formula(row_total, 6, total_formula, final_fmt)

            ws.write(row_total + 1, 1, "в том числе, работы:", total_label_fmt)
            # 🔧 FIX 1: Теперь SUMIF работает, т.к. в столбце A настоящие числа
            ws.write_formula(row_total + 1, 6, f'=SUMIF(A2:A{excel_row},">0",G2:G{excel_row})', final_fmt)

            ws.write(row_total + 2, 1, "материалы:", total_label_fmt)
            ws.write_formula(row_total + 2, 6, f'=SUMIFS(G2:G{excel_row},A2:A{excel_row},"",B2:B{excel_row},"<>*ИТОГО*",B2:B{excel_row},"<>*РАЗДЕЛ*")', final_fmt)

            if seen_works or seen_materials:
                meta_data = self.db[self.db['Работа'].isin(seen_works) | self.db['Материал'].isin(seen_materials)].drop_duplicates().reset_index(drop=True)
                if not meta_data.empty: meta_data.to_excel(writer, sheet_name="Meta", index=False)

            wb.close()
            messagebox.showinfo("Excel", f"Смета выгружена успешно!\nФайл: {export_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать файл.\nЗакройте Excel и попробуйте снова.\n\n{e}")
    
    def undo_action(self, event=None):
        """Отмена последнего изменения (Ctrl+Z)."""
        if self.undo_stack:
            item_id, col_idx, old_val = self.undo_stack.pop()
            if self.tree_smeta.exists(item_id):
                vals = list(self.tree_smeta.item(item_id, 'values'))
                vals[col_idx] = old_val
                self.tree_smeta.item(item_id, values=tuple(vals))
                self._recalculate_all_costs()

    def _show_context_menu(self, event):
        """Показывает контекстное меню ПКМ."""
        self.ctx_menu_item = self.tree_smeta.identify_row(event.y)
        self.ctx_menu_col = self.tree_smeta.identify_column(event.x)
        if not self.ctx_menu_item or not self.ctx_menu_col: return

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Копировать значение", command=self._copy_cell)
        menu.add_command(label="Вставить значение", command=self._paste_cell)
        menu.tk_popup(event.x_root, event.y_root)

    def _copy_cell(self):
        """Копирует значение выбранной ячейки в буфер."""
        if self.ctx_menu_item:
            col_idx = int(self.ctx_menu_col[1:]) - 1
            val = self.tree_smeta.item(self.ctx_menu_item, 'values')[col_idx]
            self.root.clipboard_clear()
            self.root.clipboard_append(str(val))

    def _paste_cell(self):
        """Вставляет значение из буфера и открывает редактор."""
        if self.ctx_menu_item:
            try: clip = self.root.clipboard_get().strip()
            except: return
            col_idx = int(self.ctx_menu_col[1:]) - 1
            # Запускаем встроенный редактор с уже вставленным текстом
            self._start_cell_edit(self.ctx_menu_item, self.tree_smeta.item(self.ctx_menu_item, 'values'), col_idx)
            if hasattr(self, 'edit_entry') and self.edit_entry and self.edit_entry.winfo_exists():
                self.edit_entry.delete(0, tk.END)
                self.edit_entry.insert(0, clip)
                self.edit_entry.icursor(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = SmetaApp(root)
    root.mainloop()