# -*- coding: utf-8 -*-
"""
app.py — Главный модуль приложения «Сметчик PRO 5.3» (PyQt6).
Графический интерфейс: два таба — «Составление сметы» и «Справочник».
Управление базами, черновик, настройки, экспорт/импорт.
"""
import os
import sys
import json
import math
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QMessageBox, QFileDialog, QDialog,
    QFormLayout, QGroupBox, QHeaderView, QSpinBox, QDoubleSpinBox,
    QCheckBox, QSplitter, QMenuBar, QMenu, QStatusBar, QFrame,
    QAbstractItemView, QInputDialog, QDialogButtonBox,
)
from PyQt6.QtCore import Qt, pyqtSignal, QDate
from PyQt6.QtGui import QAction, QFont, QKeySequence, QShortcut, QColor

from smeta_core import (
    build_work_block, rebuild_smeta, compute_grand_totals,
    export_smeta_to_excel, parse_exported_sheet,
    to_float, is_section, is_work, is_total, is_material, clean_name,
    SECTION_PREFIX, WORK_PREFIX, MATERIAL_PREFIX, TOTAL_PREFIX,
    CALC_HEADERS, COLS,
)
from db_manager import DatabaseManager

# ---------------------------------------------------------------------------
# Константы
# ---------------------------------------------------------------------------
DRAFT_FILE = "smeta_draft.json"
WINDOW_SIZE = (1400, 900)
FONT_FAMILY = "Segoe UI"
FONT_SIZE = 10
TABLE_ROW_HEIGHT = 28

# Индексы колонок
COL_NUM = 0
COL_NAME = 1
COL_UNIT = 2
COL_NORM1 = 3
COL_VOL1 = 4
COL_PRICE1 = 5
COL_COST1 = 6
COL_NORM2 = 7
COL_VOL2 = 8
COL_PRICE2 = 9
COL_COST2 = 10

EDITABLE_WORK = {COL_NORM1, COL_VOL1, COL_PRICE1, COL_NORM2, COL_VOL2, COL_PRICE2}
EDITABLE_MAT = {COL_NORM1, COL_PRICE1, COL_NORM2, COL_PRICE2}


def _cell(value, right_align=False):
    """Создаёт QTableWidgetItem."""
    item = QTableWidgetItem(str(value) if value is not None else "")
    if right_align:
        item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
    return item


# ---------------------------------------------------------------------------
# Диалог создания новой базы
# ---------------------------------------------------------------------------
class NewDatabaseDialog(QDialog):
    """Диалог создания новой базы данных или выбора существующей."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Новая база данных")
        self.setFixedSize(500, 300)

        layout = QVBoxLayout()

        # Выбор папки
        form = QFormLayout()
        self.folder_edit = QLineEdit()
        self.folder_edit.setText(os.path.join(os.getcwd(), "db"))
        self.folder_edit.setPlaceholderText("Папка для базы данных")

        self.name_edit = QLineEdit()
        self.name_edit.setText("smeta_db")
        self.name_edit.setPlaceholderText("Имя базы (без расширения)")

        btn_browse = QPushButton("...")
        btn_browse.clicked.connect(self._browse_folder)

        folder_row = QHBoxLayout()
        folder_row.addWidget(self.folder_edit)
        folder_row.addWidget(btn_browse)

        form.addRow("Папка базы данных:", folder_row)
        form.addRow("Имя базы данных:", self.name_edit)
        layout.addLayout(form)

        # Описание
        desc = QLabel(
            "Создаёт новую пустую базу данных SQLite в указанной папке.\n"
            "Все существующие данные в этой папке будут сохранены в резервную копию."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; font-size: 10pt;")
        layout.addWidget(desc)

        # Кнопки
        btn_row = QHBoxLayout()
        btn_create = QPushButton("Создать базу")
        btn_create.setDefault(True)
        btn_cancel = QPushButton("Отмена")

        btn_create.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)

        btn_row.addStretch()
        btn_row.addWidget(btn_create)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

        self.setLayout(layout)

    def _browse_folder(self):
        d = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if d:
            self.folder_edit.setText(d)

    def get_values(self):
        return (
            self.folder_edit.text().strip(),
            self.name_edit.text().strip(),
        )


# ---------------------------------------------------------------------------
# Диалог настроек
# ---------------------------------------------------------------------------
class SettingsDialog(QDialog):
    """Диалог настроек приложения."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.setFixedSize(450, 280)

        folder_w = QLineEdit()
        export_w = QLineEdit()
        koeff_w = QDoubleSpinBox()
        koeff_w.setRange(0.1, 3.0)
        koeff_w.setDecimals(3)
        koeff_w.setValue(1.0)

        btn_browse_folder = QPushButton("...")
        btn_browse_export = QPushButton("...")

        def on_browse_folder():
            d = QFileDialog.getExistingDirectory(self, "Папка базы данных")
            if d:
                folder_w.setText(d)

        def on_browse_export():
            d = QFileDialog.getExistingDirectory(self, "Папка экспорта смет")
            if d:
                export_w.setText(d)

        btn_browse_folder.clicked.connect(on_browse_folder)
        btn_browse_export.clicked.connect(on_browse_export)

        row_folder = QHBoxLayout()
        row_folder.addWidget(QLabel("Папка базы данных:"))
        row_folder.addWidget(folder_w)
        row_folder.addWidget(btn_browse_folder)

        row_export = QHBoxLayout()
        row_export.addWidget(QLabel("Папка экспорта:"))
        row_export.addWidget(export_w)
        row_export.addWidget(btn_browse_export)

        row_koeff = QHBoxLayout()
        row_koeff.addWidget(QLabel("Коэффициент В1 от В2:"))
        row_koeff.addWidget(koeff_w)

        btn_save = QPushButton("Сохранить")
        btn_cancel = QPushButton("Отмена")
        btn_save.setDefault(True)
        btn_save.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)

        row_btn = QHBoxLayout()
        row_btn.addWidget(btn_save)
        row_btn.addWidget(btn_cancel)

        layout = QVBoxLayout()
        layout.addLayout(row_folder)
        layout.addLayout(row_export)
        layout.addLayout(row_koeff)
        layout.addStretch()
        layout.addLayout(row_btn)
        self.setLayout(layout)

        self.folder_edit = folder_w
        self.export_edit = export_w
        self.koeff_spin = koeff_w

    def get_values(self):
        return (
            self.folder_edit.text().strip(),
            self.export_edit.text().strip(),
            self.koeff_spin.value(),
        )


# ---------------------------------------------------------------------------
# Диалог управления разделами
# ---------------------------------------------------------------------------
class SectionDialog(QDialog):
    """Диалог для управления разделами сметы."""

    def __init__(self, sections, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Разделы сметы")
        self.setMinimumSize(400, 300)
        self.sections = list(sections)

        layout = QVBoxLayout()

        form = QFormLayout()
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Название нового раздела")
        form.addRow("Название:", self.name_edit)

        btn_add = QPushButton("Добавить раздел")
        btn_add.clicked.connect(self._add_section)
        form.addRow("", btn_add)
        layout.addLayout(form)

        self.list_widget = QTableWidget()
        self.list_widget.setColumnCount(2)
        self.list_widget.setHorizontalHeaderLabels(["№", "Название"])
        self.list_widget.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.list_widget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.list_widget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._refresh_list()
        layout.addWidget(self.list_widget)

        btn_remove = QPushButton("Удалить выбранный")
        btn_remove.clicked.connect(self._remove_selected)
        btn_ok = QPushButton("OK")
        btn_cancel = QPushButton("Отмена")
        btn_ok.setDefault(True)

        row_btn = QHBoxLayout()
        row_btn.addWidget(btn_remove)
        row_btn.addStretch()
        row_btn.addWidget(btn_cancel)
        row_btn.addWidget(btn_ok)
        layout.addLayout(row_btn)

        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        self.setLayout(layout)

    def _refresh_list(self):
        self.list_widget.setRowCount(len(self.sections))
        for i, sec in enumerate(self.sections):
            self.list_widget.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.list_widget.setItem(i, 1, QTableWidgetItem(sec))

    def _add_section(self):
        name = self.name_edit.text().strip()
        if not name:
            return
        if name in self.sections:
            QMessageBox.warning(self, "Ошибка", "Раздел уже существует.")
            return
        self.sections.append(name)
        self.name_edit.clear()
        self._refresh_list()

    def _remove_selected(self):
        rows = self.list_widget.selectionModel().selectedRows()
        if not rows:
            return
        self.sections.pop(rows[0].row())
        self._refresh_list()

    def get_sections(self):
        return self.sections


# ---------------------------------------------------------------------------
# Главное окно
# ---------------------------------------------------------------------------
class SmetaMainWindow(QMainWindow):
    """Главное окно приложения «Сметчик PRO 5.2»."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Сметчик PRO 5.3")
        self.resize(*WINDOW_SIZE)

        # --- Состояние ---
        self.db_folder = os.path.join(os.getcwd(), "db")
        self.export_folder = os.path.join(os.getcwd(), "export")
        self.koeff_price = 1.0
        self.db_mgr = None

        # Данные сметы
        self.smeta_title = "Смета"
        self.smeta_rows = []
        self.extra_overhead1 = 0.0
        self.extra_overhead2 = 0.0
        self.extra_lifting1 = 0.0
        self.extra_lifting2 = 0.0
        self.extra_trash1 = 0.0
        self.extra_trash2 = 0.0
        self.sections = []
        self.works_combo_list = []

        # Undo-стек (простой список состояний)
        self._undo_stack = []
        self._redo_stack = []

        # Копируемые значения
        self.copied_value = None
        self.copied_row = None

        # --- БД ---
        os.makedirs(self.db_folder, exist_ok=True)
        self.db_mgr = DatabaseManager(self.db_folder, "smeta_db")

        # --- Меню ---
        self._create_menu()

        # --- Центральный виджет ---
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(6)

        # --- Табы ---
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # --- Вкладка: Смета ---
        self._create_smeta_tab()
        self.tabs.addTab(self.smeta_tab, "Составление сметы")

        # --- Вкладка: Справочник ---
        self._create_reference_tab()
        self.tabs.addTab(self.ref_tab, "Справочник")

        # --- Статус-бар ---
        self.statusBar().showMessage("Готово")

        # --- Горячие клавиши ---
        QShortcut(QKeySequence("Ctrl+Z"), self, self._undo)
        QShortcut(QKeySequence("Ctrl+Y"), self, self._redo)

        # --- Загружаем черновик ---
        self._load_draft()
        self._refresh_works_combo()

    # =======================================================================
    # Меню
    # =======================================================================
    def _create_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("&Файл")

        act = QAction("💾 &Сохранить черновик", self)
        act.setShortcut(QKeySequence("Ctrl+S"))
        act.triggered.connect(self._save_draft)
        file_menu.addAction(act)

        act = QAction("📤 &Экспорт в Excel", self)
        act.setShortcut(QKeySequence("Ctrl+E"))
        act.triggered.connect(self._export_to_excel)
        file_menu.addAction(act)

        act = QAction("📥 &Импорт из Excel", self)
        act.setShortcut(QKeySequence("Ctrl+I"))
        act.triggered.connect(self._import_from_excel)
        file_menu.addAction(act)

        file_menu.addSeparator()

        act = QAction("Вы&ход", self)
        act.setShortcut(QKeySequence("Ctrl+Q"))
        act.triggered.connect(self.close)
        file_menu.addAction(act)

        edit_menu = menubar.addMenu("&Правка")

        act = QAction("&Отменить (Ctrl+Z)", self)
        act.setShortcut(QKeySequence.StandardKey.Undo)
        act.triggered.connect(self._undo)
        edit_menu.addAction(act)

        act = QAction("Повторить (Ctrl+Y)", self)
        act.setShortcut(QKeySequence("Ctrl+Y"))
        act.triggered.connect(self._redo)
        edit_menu.addAction(act)

        edit_menu.addSeparator()

        act = QAction("&Добавить работу", self)
        act.setShortcut(QKeySequence("Ctrl+W"))
        act.triggered.connect(self._add_work_row)
        edit_menu.addAction(act)

        act = QAction("Добавить &раздел", self)
        act.setShortcut(QKeySequence("Ctrl+D"))
        act.triggered.connect(self._show_section_dialog)
        edit_menu.addAction(act)

        edit_menu.addSeparator()

        act = QAction("📖 &Справка", self)
        act.setShortcut(QKeySequence("F1"))
        act.triggered.connect(self._show_about)
        edit_menu.addAction(act)

        help_menu = menubar.addMenu("&Справка")
        act = QAction("&О программе", self)
        act.setShortcut(QKeySequence.StandardKey.HelpContents)
        act.triggered.connect(self._show_about)
        help_menu.addAction(act)

    # =======================================================================
    # Вкладка «Смета»
    # =======================================================================
    def _create_smeta_tab(self):
        self.smeta_tab = QWidget()
        layout = QVBoxLayout(self.smeta_tab)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        # Верхняя панель
        top_row = QHBoxLayout()
        self.title_edit = QLineEdit()
        self.title_edit.setText(self.smeta_title)
        self.title_edit.setFixedWidth(350)
        self.title_edit.textChanged.connect(self._on_title_changed)
        top_row.addWidget(QLabel("Название сметы:"))
        top_row.addWidget(self.title_edit)
        top_row.addStretch()

        self.btn_new_db = QPushButton("📁 Новая база")
        self.btn_new_db.clicked.connect(self._show_new_database_dialog)
        self.btn_new_db.setFixedWidth(130)

        self.btn_sections = QPushButton("📑 Разделы")
        self.btn_sections.clicked.connect(self._show_section_dialog)
        self.btn_sections.setFixedWidth(100)

        top_row.addWidget(self.btn_new_db)
        top_row.addWidget(self.btn_sections)
        layout.addLayout(top_row)

        # Панель добавления работы
        add_row = QHBoxLayout()
        self.work_combo = QComboBox()
        self.work_combo.setFixedWidth(350)
        self.work_combo.setEditable(True)
        self.work_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.work_combo.currentTextChanged.connect(self._on_work_combo_changed)

        btn_add = QPushButton("➕ Добавить")
        btn_add.clicked.connect(self._add_work_row)
        btn_add.setFixedWidth(110)

        self.vol_spin = QDoubleSpinBox()
        self.vol_spin.setRange(0.001, 999999)
        self.vol_spin.setDecimals(3)
        self.vol_spin.setValue(1.0)
        self.vol_spin.setFixedWidth(100)

        self.price1_spin = QDoubleSpinBox()
        self.price1_spin.setRange(0, 99999999)
        self.price1_spin.setDecimals(2)
        self.price1_spin.setFixedWidth(120)

        self.price2_spin = QDoubleSpinBox()
        self.price2_spin.setRange(0, 99999999)
        self.price2_spin.setDecimals(2)
        self.price2_spin.setFixedWidth(120)

        add_row.addWidget(QLabel("Добавить работу:"))
        add_row.addWidget(self.work_combo)
        add_row.addWidget(btn_add)
        add_row.addWidget(QLabel("Объём:"))
        add_row.addWidget(self.vol_spin)
        add_row.addWidget(QLabel("Цена В1:"))
        add_row.addWidget(self.price1_spin)
        add_row.addWidget(QLabel("Цена В2:"))
        add_row.addWidget(self.price2_spin)
        layout.addLayout(add_row)

        # Таблица сметы
        self.smeta_table = QTableWidget()
        self.smeta_table.setColumnCount(len(CALC_HEADERS))
        self.smeta_table.setHorizontalHeaderLabels(CALC_HEADERS)
        self.smeta_table.setRowCount(0)
        self.smeta_table.verticalHeader().setVisible(False)
        self.smeta_table.setAlternatingRowColors(True)
        self.smeta_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.smeta_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.smeta_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.smeta_table.verticalHeader().setDefaultSectionSize(TABLE_ROW_HEIGHT)
        self.smeta_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.smeta_table.horizontalHeader().setStretchLastSection(True)
        self.smeta_table.cellDoubleClicked.connect(self._on_cell_double_click)
        self.smeta_table.customContextMenuRequested.connect(self._show_table_context_menu)
        layout.addWidget(self.smeta_table)

        # Итоги
        totals_widget = QWidget()
        totals_layout = QHBoxLayout(totals_widget)
        self.lbl_total1 = QLabel("Итого В1: 0.00 руб.")
        self.lbl_total1.setFont(QFont(FONT_FAMILY, 11, QFont.Weight.Bold))
        self.lbl_total2 = QLabel("Итого В2: 0.00 руб.")
        self.lbl_total2.setFont(QFont(FONT_FAMILY, 11, QFont.Weight.Bold))
        self.lbl_savings = QLabel("Экономия: 0.00 руб. (0.0%)")
        self.lbl_savings.setFont(QFont(FONT_FAMILY, 10))
        totals_layout.addWidget(self.lbl_total1)
        totals_layout.addWidget(self.lbl_total2)
        totals_layout.addWidget(self.lbl_savings)
        totals_layout.addStretch()
        layout.addWidget(totals_widget)

        # Доп. расходы
        exp_widget = QWidget()
        exp_layout = QFormLayout(exp_widget)
        self.overhead1_spin = QDoubleSpinBox(); self.overhead1_spin.setRange(0, 99999999); self.overhead1_spin.setDecimals(2)
        self.overhead2_spin = QDoubleSpinBox(); self.overhead2_spin.setRange(0, 99999999); self.overhead2_spin.setDecimals(2)
        self.lifting1_spin = QDoubleSpinBox(); self.lifting1_spin.setRange(0, 99999999); self.lifting1_spin.setDecimals(2)
        self.lifting2_spin = QDoubleSpinBox(); self.lifting2_spin.setRange(0, 99999999); self.lifting2_spin.setDecimals(2)
        self.trash1_spin = QDoubleSpinBox(); self.trash1_spin.setRange(0, 99999999); self.trash1_spin.setDecimals(2)
        self.trash2_spin = QDoubleSpinBox(); self.trash2_spin.setRange(0, 99999999); self.trash2_spin.setDecimals(2)

        exp_layout.addRow("Накладные (В1):", self.overhead1_spin)
        exp_layout.addRow("Накладные (В2):", self.overhead2_spin)
        exp_layout.addRow("Подъёмные (В1):", self.lifting1_spin)
        exp_layout.addRow("Подъёмные (В2):", self.lifting2_spin)
        exp_layout.addRow("Вывоз мусора (В1):", self.trash1_spin)
        exp_layout.addRow("Вывоз мусора (В2):", self.trash2_spin)
        layout.addWidget(exp_widget)

        # Кнопки
        btn_row = QHBoxLayout()
        btn_save = QPushButton("💾 Сохранить"); btn_save.clicked.connect(self._save_draft)
        btn_export = QPushButton("📤 Экспорт Excel"); btn_export.clicked.connect(self._export_to_excel)
        btn_import = QPushButton("📥 Импорт Excel"); btn_import.clicked.connect(self._import_from_excel)
        btn_recalc = QPushButton("🔄 Пересчитать"); btn_recalc.clicked.connect(self._recalc_smeta)
        btn_clear = QPushButton("🗑 Очистить смету"); btn_clear.clicked.connect(self._clear_smeta)
        btn_settings = QPushButton("⚙️ Настройки")
        btn_settings.clicked.connect(self.show_settings)

        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_export)
        btn_row.addWidget(btn_import)
        btn_row.addWidget(btn_recalc)
        btn_row.addStretch()
        btn_row.addWidget(btn_settings)
        btn_row.addWidget(btn_clear)
        layout.addLayout(btn_row)

        # Сигналы spinbox-ов
        for sp in [self.overhead1_spin, self.overhead2_spin,
                   self.lifting1_spin, self.lifting2_spin,
                   self.trash1_spin, self.trash2_spin]:
            sp.valueChanged.connect(self._on_extra_changed)

    # =======================================================================
    # Вкладка «Справочник»
    # =======================================================================
    def _create_reference_tab(self):
        self.ref_tab = QWidget()
        layout = QVBoxLayout(self.ref_tab)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        self.ref_tabs = QTabWidget()
        layout.addWidget(self.ref_tabs)

        self._create_works_table_tab()
        self.ref_tabs.addTab(self.works_tab, "Работы")

        self._create_materials_table_tab()
        self.ref_tabs.addTab(self.materials_tab, "Материалы")

        self._create_links_tab()
        self.ref_tabs.addTab(self.links_tab, "Связи")

    def _create_works_table_tab(self):
        self.works_tab = QWidget()
        layout = QVBoxLayout(self.works_tab)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        filter_row = QHBoxLayout()
        self.works_filter_edit = QLineEdit()
        self.works_filter_edit.setPlaceholderText("Поиск...")
        self.works_filter_edit.setFixedWidth(300)
        self.works_filter_edit.textChanged.connect(self._filter_works)
        filter_row.addWidget(QLabel("Поиск:"))
        filter_row.addWidget(self.works_filter_edit)
        filter_row.addStretch()
        layout.addLayout(filter_row)

        self.works_table = QTableWidget()
        self.works_table.setColumnCount(5)
        self.works_table.setHorizontalHeaderLabels(["№", "Название", "Ед. изм.", "Цена В1", "Цена В2"])
        self.works_table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self.works_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.works_table.verticalHeader().setVisible(False)
        self.works_table.verticalHeader().setDefaultSectionSize(TABLE_ROW_HEIGHT)
        self.works_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.works_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.works_table)

        form = QFormLayout()
        self.ref_work_name = QLineEdit()
        self.ref_work_unit = QLineEdit()
        self.ref_work_price1 = QDoubleSpinBox(); self.ref_work_price1.setRange(0, 99999999); self.ref_work_price1.setDecimals(2)
        self.ref_work_price2 = QDoubleSpinBox(); self.ref_work_price2.setRange(0, 99999999); self.ref_work_price2.setDecimals(2)
        form.addRow("Название:", self.ref_work_name)
        form.addRow("Ед. изм.:", self.ref_work_unit)
        form.addRow("Цена В1:", self.ref_work_price1)
        form.addRow("Цена В2:", self.ref_work_price2)
        layout.addLayout(form)

        btn_row = QHBoxLayout()
        self.btn_add_work = QPushButton("➕ Добавить"); self.btn_add_work.clicked.connect(self._add_work)
        self.btn_update_work = QPushButton("✏️ Обновить"); self.btn_update_work.clicked.connect(self._update_work)
        self.btn_delete_work = QPushButton("🗑 Удалить"); self.btn_delete_work.clicked.connect(self._delete_work)
        btn_row.addWidget(self.btn_add_work)
        btn_row.addWidget(self.btn_update_work)
        btn_row.addWidget(self.btn_delete_work)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._refresh_works_table()

    def _create_materials_table_tab(self):
        self.materials_tab = QWidget()
        layout = QVBoxLayout(self.materials_tab)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        filter_row = QHBoxLayout()
        self.materials_filter_edit = QLineEdit()
        self.materials_filter_edit.setPlaceholderText("Поиск...")
        self.materials_filter_edit.setFixedWidth(300)
        self.materials_filter_edit.textChanged.connect(self._filter_materials)
        filter_row.addWidget(QLabel("Поиск:"))
        filter_row.addWidget(self.materials_filter_edit)
        filter_row.addStretch()
        layout.addLayout(filter_row)

        self.materials_table = QTableWidget()
        self.materials_table.setColumnCount(5)
        self.materials_table.setHorizontalHeaderLabels(["№", "Название", "Ед. изм.", "Цена В1", "Цена В2"])
        self.materials_table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self.materials_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.materials_table.verticalHeader().setVisible(False)
        self.materials_table.verticalHeader().setDefaultSectionSize(TABLE_ROW_HEIGHT)
        self.materials_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.materials_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.materials_table)

        form = QFormLayout()
        self.ref_mat_name = QLineEdit()
        self.ref_mat_unit = QLineEdit()
        self.ref_mat_price1 = QDoubleSpinBox(); self.ref_mat_price1.setRange(0, 99999999); self.ref_mat_price1.setDecimals(2)
        self.ref_mat_price2 = QDoubleSpinBox(); self.ref_mat_price2.setRange(0, 99999999); self.ref_mat_price2.setDecimals(2)
        form.addRow("Название:", self.ref_mat_name)
        form.addRow("Ед. изм.:", self.ref_mat_unit)
        form.addRow("Цена В1:", self.ref_mat_price1)
        form.addRow("Цена В2:", self.ref_mat_price2)
        layout.addLayout(form)

        btn_row = QHBoxLayout()
        self.btn_add_mat = QPushButton("➕ Добавить"); self.btn_add_mat.clicked.connect(self._add_material)
        self.btn_update_mat = QPushButton("✏️ Обновить"); self.btn_update_mat.clicked.connect(self._update_material)
        self.btn_delete_mat = QPushButton("🗑 Удалить"); self.btn_delete_mat.clicked.connect(self._delete_material)
        btn_row.addWidget(self.btn_add_mat)
        btn_row.addWidget(self.btn_update_mat)
        btn_row.addWidget(self.btn_delete_mat)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._refresh_materials_table()

    def _create_links_tab(self):
        self.links_tab = QWidget()
        layout = QVBoxLayout(self.links_tab)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        combo_row = QHBoxLayout()
        self.link_work_combo = QComboBox()
        self.link_work_combo.setEditable(True)
        self.link_work_combo.setFixedWidth(400)
        self.link_work_combo.currentTextChanged.connect(self._refresh_links_table)
        combo_row.addWidget(QLabel("Работа:"))
        combo_row.addWidget(self.link_work_combo)
        combo_row.addStretch()
        layout.addLayout(combo_row)

        self.links_table = QTableWidget()
        self.links_table.setColumnCount(6)
        self.links_table.setHorizontalHeaderLabels(["Материал", "Ед. изм.", "Расход В1", "Расход В2", "Цена В1", "Цена В2"])
        self.links_table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self.links_table.verticalHeader().setVisible(False)
        self.links_table.verticalHeader().setDefaultSectionSize(TABLE_ROW_HEIGHT)
        self.links_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.links_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.links_table)

        btn_row = QHBoxLayout()
        self.btn_add_link = QPushButton("➕ Добавить связь"); self.btn_add_link.clicked.connect(self._add_link)
        self.btn_remove_link = QPushButton("🗑 Удалить связь"); self.btn_remove_link.clicked.connect(self._remove_link)
        btn_row.addWidget(self.btn_add_link)
        btn_row.addWidget(self.btn_remove_link)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._refresh_link_works_combo()

    # =======================================================================
    # РАБОТЫ: refresh, filter, CRUD
    # =======================================================================
    def _refresh_works_table(self):
        df = self.db_mgr.get_works()
        self.works_table.setRowCount(len(df))
        for i, row in df.iterrows():
            self.works_table.setItem(i, 0, QTableWidgetItem(str(int(row['id']))))
            self.works_table.setItem(i, 1, QTableWidgetItem(str(row['name'])))
            self.works_table.setItem(i, 2, QTableWidgetItem(str(row['unit'])))
            self.works_table.setItem(i, 3, _cell(row['price_1'], True))
            self.works_table.setItem(i, 4, _cell(row['price_2'], True))
        self.works_table.cellDoubleClicked.connect(self._on_works_double_click)

    def _filter_works(self):
        q = self.works_filter_edit.text().strip().lower()
        for i in range(self.works_table.rowCount()):
            item = self.works_table.item(i, 1)
            if item:
                self.works_table.setRowHidden(i, bool(q and q not in item.text().lower()))

    def _on_works_double_click(self, row, col):
        df = self.db_mgr.get_works()
        if df.empty:
            return
        idx = df.index[row] if row < len(df) else 0
        rd = df.iloc[idx]
        self.ref_work_name.setText(str(rd['name']))
        self.ref_work_unit.setText(str(rd['unit']))
        self.ref_work_price1.setValue(float(rd['price_1']))
        self.ref_work_price2.setValue(float(rd['price_2']))

    def _add_work(self):
        name = self.ref_work_name.text().strip()
        unit = self.ref_work_unit.text().strip()
        p1 = self.ref_work_price1.value()
        p2 = self.ref_work_price2.value()
        if not name:
            QMessageBox.warning(self, "Ошибка", "Введите название работы.")
            return
        self.db_mgr.add_work(name, unit, p1, p2)
        self._refresh_works_table()
        self._refresh_works_combo()
        self._refresh_link_works_combo()
        self.statusBar().showMessage(f"Добавлена работа: {name}")
        self._clear_ref_work_fields()

    def _update_work(self):
        row = self.works_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите работу.")
            return
        df = self.db_mgr.get_works()
        if df.empty or row >= len(df):
            return
        wid = int(df.iloc[row]['id'])
        name = self.ref_work_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Ошибка", "Введите название работы.")
            return
        self.db_mgr.update_work(wid, name=name, unit=self.ref_work_unit.text().strip(),
                                price_1=self.ref_work_price1.value(), price_2=self.ref_work_price2.value())
        self._refresh_works_table()
        self._refresh_works_combo()
        self._refresh_link_works_combo()
        self.statusBar().showMessage(f"Обновлена работа: {name}")

    def _delete_work(self):
        row = self.works_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите работу.")
            return
        df = self.db_mgr.get_works()
        if df.empty or row >= len(df):
            return
        name = df.iloc[row]['name']
        wid = int(df.iloc[row]['id'])
        if QMessageBox.question(self, "Подтверждение",
                                 f'Удалить работу "{name}" и все привязанные материалы?',
                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                 QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        self.db_mgr.delete_work(wid)
        self._refresh_works_table()
        self._refresh_works_combo()
        self._refresh_link_works_combo()
        self.statusBar().showMessage(f"Удалена работа: {name}")

    def _clear_ref_work_fields(self):
        self.ref_work_name.clear()
        self.ref_work_unit.clear()
        self.ref_work_price1.setValue(0)
        self.ref_work_price2.setValue(0)

    # =======================================================================
    # МАТЕРИАЛЫ: refresh, filter, CRUD
    # =======================================================================
    def _refresh_materials_table(self):
        df = self.db_mgr.get_materials()
        self.materials_table.setRowCount(len(df))
        for i, row in df.iterrows():
            self.materials_table.setItem(i, 0, QTableWidgetItem(str(int(row['id']))))
            self.materials_table.setItem(i, 1, QTableWidgetItem(str(row['name'])))
            self.materials_table.setItem(i, 2, QTableWidgetItem(str(row['unit'])))
            self.materials_table.setItem(i, 3, _cell(row['price_1'], True))
            self.materials_table.setItem(i, 4, _cell(row['price_2'], True))
        self.materials_table.cellDoubleClicked.connect(self._on_materials_double_click)

    def _filter_materials(self):
        q = self.materials_filter_edit.text().strip().lower()
        for i in range(self.materials_table.rowCount()):
            item = self.materials_table.item(i, 1)
            if item:
                self.materials_table.setRowHidden(i, bool(q and q not in item.text().lower()))

    def _on_materials_double_click(self, row, col):
        df = self.db_mgr.get_materials()
        if df.empty:
            return
        idx = df.index[row] if row < len(df) else 0
        rd = df.iloc[idx]
        self.ref_mat_name.setText(str(rd['name']))
        self.ref_mat_unit.setText(str(rd['unit']))
        self.ref_mat_price1.setValue(float(rd['price_1']))
        self.ref_mat_price2.setValue(float(rd['price_2']))

    def _add_material(self):
        name = self.ref_mat_name.text().strip()
        unit = self.ref_mat_unit.text().strip()
        p1 = self.ref_mat_price1.value()
        p2 = self.ref_mat_price2.value()
        if not name:
            QMessageBox.warning(self, "Ошибка", "Введите название материала.")
            return
        self.db_mgr.add_material(name, unit, p1, p2)
        self._refresh_materials_table()
        self.statusBar().showMessage(f"Добавлен материал: {name}")
        self._clear_ref_mat_fields()

    def _update_material(self):
        row = self.materials_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите материал.")
            return
        df = self.db_mgr.get_materials()
        if df.empty or row >= len(df):
            return
        mid = int(df.iloc[row]['id'])
        name = self.ref_mat_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Ошибка", "Введите название материала.")
            return
        self.db_mgr.update_material(mid, name=name, unit=self.ref_mat_unit.text().strip(),
                                     price_1=self.ref_mat_price1.value(), price_2=self.ref_mat_price2.value())
        self._refresh_materials_table()
        self.statusBar().showMessage(f"Обновлён материал: {name}")

    def _delete_material(self):
        row = self.materials_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите материал.")
            return
        df = self.db_mgr.get_materials()
        if df.empty or row >= len(df):
            return
        name = df.iloc[row]['name']
        mid = int(df.iloc[row]['id'])
        if QMessageBox.question(self, "Подтверждение",
                                 f'Удалить материал "{name}"?',
                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                 QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        self.db_mgr.delete_material(mid)
        self._refresh_materials_table()
        self.statusBar().showMessage(f"Удалён материал: {name}")

    def _clear_ref_mat_fields(self):
        self.ref_mat_name.clear()
        self.ref_mat_unit.clear()
        self.ref_mat_price1.setValue(0)
        self.ref_mat_price2.setValue(0)

    # =======================================================================
    # СВЯЗИ
    # =======================================================================
    def _refresh_link_works_combo(self):
        df = self.db_mgr.get_works()
        names = df['name'].tolist() if not df.empty else []
        cur = self.link_work_combo.currentText()
        self.link_work_combo.blockSignals(True)
        self.link_work_combo.clear()
        self.link_work_combo.addItems(names)
        if cur and cur in names:
            self.link_work_combo.setCurrentText(cur)
        elif names:
            self.link_work_combo.setCurrentIndex(0)
        self.link_work_combo.blockSignals(False)

    def _refresh_links_table(self):
        wn = self.link_work_combo.currentText().strip()
        if not wn:
            self.links_table.setRowCount(0)
            return
        wd = self.db_mgr.get_work_with_materials(wn)
        if wd is None:
            self.links_table.setRowCount(0)
            return
        mats = wd['materials']
        self.links_table.setRowCount(len(mats))
        for i, m in enumerate(mats):
            self.links_table.setItem(i, 0, QTableWidgetItem(m['name']))
            self.links_table.setItem(i, 1, QTableWidgetItem(m['unit']))
            self.links_table.setItem(i, 2, _cell(m['consumption_1'], True))
            self.links_table.setItem(i, 3, _cell(m['consumption_2'], True))
            self.links_table.setItem(i, 4, _cell(m['price_1'], True))
            self.links_table.setItem(i, 5, _cell(m['price_2'], True))
        self.links_table.cellDoubleClicked.connect(self._on_link_double_click)

    def _on_link_double_click(self, row, col):
        if col not in (2, 3):
            return
        cur = self.links_table.item(row, col).text()
        val, ok = QInputDialog.getDouble(self, "Изменение расхода", "Введите новое значение:",
                                          float(cur) if cur and cur != "-" else 0.0,
                                          0.0, 999999, 3)
        if ok:
            self.links_table.setItem(row, col, _cell(val, True))

    def _add_link(self):
        wn = self.link_work_combo.currentText().strip()
        if not wn:
            QMessageBox.warning(self, "Ошибка", "Выберите работу.")
            return
        df = self.db_mgr.get_materials()
        if df.empty:
            QMessageBox.warning(self, "Ошибка", "Нет материалов в справочнике.")
            return
        mat_names = df['name'].tolist()
        mat_name, ok = QInputDialog.getItem(self, "Выбор материала", "Материал:", mat_names, 0, False)
        if not ok:
            return
        work = self.db_mgr.get_work_by_name(wn)
        mat = self.db_mgr.get_material_by_name(mat_name)
        if work is None or mat is None:
            QMessageBox.critical(self, "Ошибка", "Не найдена работа или материал.")
            return
        c1, ok1 = QInputDialog.getDouble(self, "Расход В1", "Норма расхода В1:", 0.0, 0, 999999, 3)
        if not ok1:
            return
        c2, ok2 = QInputDialog.getDouble(self, "Расход В2", "Норма расхода В2:", 0.0, 0, 999999, 3)
        if not ok2:
            return
        self.db_mgr.add_work_material_link(int(work['id']), int(mat['id']), c1, c2)
        self._refresh_links_table()
        self.statusBar().showMessage(f"Связь: {wn} → {mat_name}")

    def _remove_link(self):
        row = self.links_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Ошибка", "Выберите связь.")
            return
        wn = self.link_work_combo.currentText().strip()
        mn = self.links_table.item(row, 0).text() if self.links_table.item(row, 0) else ""
        if not wn or not mn:
            return
        work = self.db_mgr.get_work_by_name(wn)
        mat = self.db_mgr.get_material_by_name(mn)
        if work is None or mat is None:
            return
        self.db_mgr.remove_work_material_link(int(work['id']), int(mat['id']))
        self._refresh_links_table()
        self.statusBar().showMessage(f"Связь удалена: {wn} → {mn}")

    # =======================================================================
    # ComboBox
    # =======================================================================
    def _refresh_works_combo(self):
        df = self.db_mgr.get_works()
        names = df['name'].tolist() if not df.empty else []
        cur = self.work_combo.currentText()
        self.work_combo.blockSignals(True)
        self.work_combo.clear()
        self.work_combo.addItems(names)
        if cur and cur in names:
            self.work_combo.setCurrentText(cur)
        elif names:
            self.work_combo.setCurrentIndex(0)
        self.work_combo.blockSignals(False)

    def _on_work_combo_changed(self, text):
        if not text:
            return
        work = self.db_mgr.get_work_by_name(text)
        if work is not None:
            self.price1_spin.setValue(float(work['price_1']))
            self.price2_spin.setValue(float(work['price_2']))

    # =======================================================================
    # Действия со сметой
    # =======================================================================
    def _add_work_row(self):
        work_name = self.work_combo.currentText().strip()
        if not work_name:
            QMessageBox.warning(self, "Ошибка", "Выберите или введите название работы.")
            return
        vol = self.vol_spin.value()
        block = build_work_block(work_name, vol, None, len(self.smeta_rows) + 1, self.db_mgr)
        if block is None:
            QMessageBox.warning(self, "Ошибка",
                                f'Работа "{work_name}" не найдена в справочнике.\n'
                                f'Сначала добавьте её во вкладку "Справочник → Работы".')
            return
        self._push_undo_state()
        self.smeta_rows.extend(block)
        self._save_draft()
        self._update_table_from_rows()
        self.statusBar().showMessage(f"Добавлена: {work_name} (объём: {vol})")

    def _show_section_dialog(self):
        dlg = SectionDialog(self.sections, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._push_undo_state()
            self.sections = dlg.get_sections()
            self._rebuild_with_sections()
            self._save_draft()
            self._update_table_from_rows()

    def _show_new_database_dialog(self):
        """Показывает диалог создания новой базы данных."""
        dlg = NewDatabaseDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            new_folder, new_name = dlg.get_values()
            if not new_folder or not new_name:
                QMessageBox.warning(self, "Ошибка", "Укажите папку и имя базы данных.")
                return

            # Подтверждение
            reply = QMessageBox.question(
                self, "Подтверждение",
                f'Создать новую базу "{new_name}" в папке:\n{new_folder}\n\n'
                f"Текущая база будет закрыта. Черновик будет сохранён.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

            # Создаём новую базу
            try:
                os.makedirs(new_folder, exist_ok=True)
                self.db_mgr.flush()
                self.db_mgr.close()

                self.db_folder = new_folder
                self.db_mgr = DatabaseManager(new_folder, new_name)

                self.statusBar().showMessage(f"База создана: {os.path.join(new_folder, new_name + '.db')}")
                QMessageBox.information(self, "База создана",
                                       f"Новая база данных создана:\n{os.path.join(new_folder, new_name + '.db')}")

                self._refresh_works_combo()
            except Exception as e:
                QMessageBox.critical(self, "Ошибка создания базы", str(e))

    def _rebuild_with_sections(self):
        filtered = [r for r in self.smeta_rows if not is_section(str(r[1]))]
        new_rows = []
        for sec in self.sections:
            new_rows.append(("", f"{SECTION_PREFIX}{sec}", "", "", "", "", "", "", "", "", ""))
        new_rows.extend(filtered)
        self.smeta_rows = new_rows

    def _add_expense_row(self):
        self.smeta_rows.append(("Накладные и транспортные расходы", "", "", "",
                                self.extra_overhead1, "", "", "",
                                self.extra_overhead2, "", ""))
        self._save_draft()
        self._update_table_from_rows()

    def _recalc_smeta(self):
        self._push_undo_state()
        self.smeta_rows = rebuild_smeta(self.smeta_rows)
        self._update_table_from_rows()
        self._save_draft()
        self.statusBar().showMessage("Смета пересчитана")

    def _clear_smeta(self):
        if QMessageBox.question(self, "Подтверждение", "Очистить всю смету?",
                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                 QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        self._push_undo_state()
        self.smeta_rows = []
        self.sections = []
        self.smeta_title = "Смета"
        self.title_edit.setText(self.smeta_title)
        self._update_table_from_rows()
        self._update_totals()
        self._save_draft()
        self.statusBar().showMessage("Смета очищена")

    # =======================================================================
    # Обновление таблицы и итогов
    # =======================================================================
    def _update_table_from_rows(self):
        self.smeta_table.setRowCount(len(self.smeta_rows))
        for r_idx, row in enumerate(self.smeta_rows):
            for c_idx in range(len(CALC_HEADERS)):
                val = row[c_idx] if c_idx < len(row) else ""
                ra = c_idx in (COL_NUM, COL_COST1, COL_COST2, COL_NORM1, COL_NORM2,
                               COL_VOL1, COL_VOL2, COL_PRICE1, COL_PRICE2)
                item = _cell(val, ra)
                item.setData(Qt.ItemDataRole.UserRole, list(row))
                self.smeta_table.setItem(r_idx, c_idx, item)

            name = str(row[1]).strip() if len(row) > 1 else ""
            if is_section(name):
                bg = QColor(225, 190, 231)
                for c in range(len(CALC_HEADERS)):
                    it = self.smeta_table.item(r_idx, c)
                    if it: it.setBackground(bg)
            elif is_work(name):
                bg = QColor(255, 242, 204)
                for c in range(len(CALC_HEADERS)):
                    it = self.smeta_table.item(r_idx, c)
                    if it:
                        it.setBackground(bg)
                        it.setFont(QFont(FONT_FAMILY, FONT_SIZE, QFont.Weight.Bold))
            elif is_material(name):
                for c in range(len(CALC_HEADERS)):
                    it = self.smeta_table.item(r_idx, c)
                    if it:
                        it.setFont(QFont(FONT_FAMILY, FONT_SIZE, QFont.Weight.Normal))
            elif is_total(name):
                bg = QColor(189, 215, 238)
                for c in range(len(CALC_HEADERS)):
                    it = self.smeta_table.item(r_idx, c)
                    if it:
                        it.setBackground(bg)
                        it.setFont(QFont(FONT_FAMILY, FONT_SIZE, QFont.Weight.Bold))

        self._update_totals()

    def _update_totals(self):
        t1, t2 = compute_grand_totals(self.smeta_rows)
        savings = t1 - t2
        ratio = (t2 / t1 * 100) if t1 > 0 else 0.0
        self.lbl_total1.setText(f"Итого В1: {t1:,.2f} руб.")
        self.lbl_total2.setText(f"Итого В2: {t2:,.2f} руб.")
        self.lbl_savings.setText(f"Экономия: {savings:,.2f} руб. ({ratio:.1f}%)")

    # =======================================================================
    # Редактирование ячеек
    # =======================================================================
    def _on_cell_double_click(self, row, col):
        if row < 0 or row >= len(self.smeta_rows):
            return
        row_data = list(self.smeta_rows[row])
        name = str(row_data[1]).strip()

        if is_section(name):
            new_name, ok = QInputDialog.getText(self, "Редактирование раздела",
                                                 "Название раздела:",
                                                 Qt.InputMode.TextInput,
                                                 clean_name(name))
            if ok and new_name:
                self._push_undo_state()
                idx = self.sections.index(clean_name(name)) if clean_name(name) in self.sections else -1
                if idx >= 0:
                    self.sections[idx] = new_name
                row_data[1] = f"{SECTION_PREFIX}{new_name}"
                self.smeta_rows[row] = tuple(row_data)
                self._update_table_from_rows()
                self._save_draft()
            return

        editable = EDITABLE_WORK if is_work(name) else (EDITABLE_MAT if is_material(name) else set())
        if not editable or col not in editable:
            return

        cur = row_data[col]
        cur_str = str(cur).replace(",", ".") if cur else "0"
        try:
            cur_num = float(cur_str)
        except (ValueError, TypeError):
            cur_num = 0.0

        decimals = 3 if col in (COL_NORM1, COL_NORM2, COL_VOL1, COL_VOL2) else 2
        val, ok = QInputDialog.getDouble(self, "Редактирование",
                                          f"Значение ({CALC_HEADERS[col]}):",
                                          cur_num, 0.0, 99999999, decimals)
        if ok:
            self._push_undo_state()
            row_data[col] = round(val, 3) if col in (COL_NORM1, COL_NORM2, COL_VOL1, COL_VOL2) else round(val, 2)
            self.smeta_rows[row] = tuple(row_data)
            self.smeta_rows = rebuild_smeta(self.smeta_rows)
            self._update_table_from_rows()
            self._save_draft()
            self.statusBar().showMessage(f"Изменено: {CALC_HEADERS[col]} = {val}")

    # =======================================================================
    # Контекстное меню
    # =======================================================================
    def _show_table_context_menu(self, pos):
        menu = self.smeta_table.createStandardContextMenu()
        menu.addSeparator()
        act_copy = QAction("📋 Копировать", self)
        act_copy.triggered.connect(self._copy_cell)
        menu.addAction(act_copy)
        act_paste = QAction("📌 Вставить", self)
        act_paste.triggered.connect(self._paste_cell)
        menu.addAction(act_paste)
        menu.exec(self.smeta_table.viewport().mapToGlobal(pos))

    def _copy_cell(self):
        sel = self.smeta_table.selectedItems()
        if not sel:
            return
        r, c = sel[0].row(), sel[0].column()
        if r < 0 or r >= len(self.smeta_rows):
            return
        self.copied_value = self.smeta_rows[r][c] if c < len(self.smeta_rows[r]) else None
        self.copied_row = r
        self.statusBar().showMessage("Значение скопировано")

    def _paste_cell(self):
        if self.copied_value is None:
            return
        sel = self.smeta_table.selectedItems()
        if not sel:
            return
        r, c = sel[0].row(), sel[0].column()
        if r < 0 or r >= len(self.smeta_rows) or c >= len(CALC_HEADERS):
            return
        self._push_undo_state()
        row_data = list(self.smeta_rows[r])
        row_data[c] = self.copied_value
        self.smeta_rows[r] = tuple(row_data)
        self.smeta_rows = rebuild_smeta(self.smeta_rows)
        self._update_table_from_rows()
        self._save_draft()

    # =======================================================================
    # Undo / Redo
    # =======================================================================
    def _push_undo_state(self):
        """Сохраняет текущее состояние в undo-стек."""
        state = {
            'rows': [list(r) for r in self.smeta_rows],
            'title': self.smeta_title,
            'sections': list(self.sections),
            'extra': (self.extra_overhead1, self.extra_overhead2,
                      self.extra_lifting1, self.extra_lifting2,
                      self.extra_trash1, self.extra_trash2),
        }
        self._undo_stack.append(state)
        self._redo_stack.clear()  # Очищаем redo-стек при новом действии

    def _undo(self):
        if not self._undo_stack:
            self.statusBar().showMessage("Нечего отменять")
            return
        state = self._undo_stack.pop()
        self._redo_stack.append({
            'rows': [list(r) for r in self.smeta_rows],
            'title': self.smeta_title,
            'sections': list(self.sections),
            'extra': (self.extra_overhead1, self.extra_overhead2,
                      self.extra_lifting1, self.extra_lifting2,
                      self.extra_trash1, self.extra_trash2),
        })
        self._restore_state(state)
        self.statusBar().showMessage("Действие отменено (Ctrl+Z)")

    def _redo(self):
        if not self._redo_stack:
            self.statusBar().showMessage("Нечего повторять")
            return
        state = self._redo_stack.pop()
        self._undo_stack.append({
            'rows': [list(r) for r in self.smeta_rows],
            'title': self.smeta_title,
            'sections': list(self.sections),
            'extra': (self.extra_overhead1, self.extra_overhead2,
                      self.extra_lifting1, self.extra_lifting2,
                      self.extra_trash1, self.extra_trash2),
        })
        self._restore_state(state)
        self.statusBar().showMessage("Действие повторено (Ctrl+Y)")

    def _restore_state(self, state):
        self.smeta_rows = [tuple(r) for r in state['rows']]
        self.smeta_title = state['title']
        self.sections = state['sections']
        self.title_edit.setText(self.smeta_title)
        extra = state['extra']
        self.extra_overhead1, self.extra_overhead2 = extra[0], extra[1]
        self.extra_lifting1, self.extra_lifting2 = extra[2], extra[3]
        self.extra_trash1, self.extra_trash2 = extra[4], extra[5]
        self.overhead1_spin.setValue(self.extra_overhead1)
        self.overhead2_spin.setValue(self.extra_overhead2)
        self.lifting1_spin.setValue(self.extra_lifting1)
        self.lifting2_spin.setValue(self.extra_lifting2)
        self.trash1_spin.setValue(self.extra_trash1)
        self.trash2_spin.setValue(self.extra_trash2)
        self._update_table_from_rows()

    # =======================================================================
    # Черновик
    # =======================================================================
    def _save_draft(self):
        draft_path = os.path.join(self.db_folder, DRAFT_FILE)
        try:
            prev_state = {}
            if os.path.exists(draft_path):
                with open(draft_path, 'r', encoding='utf-8') as f:
                    prev_state = json.load(f)

            current_state = {
                'rows': [list(r) for r in self.smeta_rows],
                'title': self.smeta_title,
                'sections': list(self.sections),
                'extra': (self.extra_overhead1, self.extra_overhead2,
                          self.extra_lifting1, self.extra_lifting2,
                          self.extra_trash1, self.extra_trash2),
            }
            draft = {
                'prev_state': prev_state,
                'current_state': current_state,
                'timestamp': datetime.now().isoformat(),
            }
            with open(draft_path, 'w', encoding='utf-8') as f:
                json.dump(draft, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.statusBar().showMessage(f"Ошибка сохранения черновика: {e}")

    def _load_draft(self):
        draft_path = os.path.join(self.db_folder, DRAFT_FILE)
        if not os.path.exists(draft_path):
            return
        try:
            with open(draft_path, 'r', encoding='utf-8') as f:
                draft = json.load(f)
            current = draft.get('current_state')
            if current is None:
                return
            self._restore_state(current)
            self.statusBar().showMessage("Черновик восстановлен")
        except Exception as e:
            self.statusBar().showMessage(f"Ошибка загрузки черновика: {e}")

    # =======================================================================
    # Экспорт / Импорт
    # =======================================================================
    def _export_to_excel(self):
        if not self.smeta_rows:
            QMessageBox.warning(self, "Ошибка", "Смета пуста.")
            return
        os.makedirs(self.export_folder, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{self.smeta_title}_{ts}.xlsx"
        path = os.path.join(self.export_folder, filename)
        try:
            meta_rows = []
            for row in self.smeta_rows:
                if is_work(row[1]) or is_material(row[1]):
                    meta_rows.append(list(row))
            export_smeta_to_excel(
                rows=self.smeta_rows, output_path=path,
                title=self.smeta_title, meta_rows=meta_rows,
                overhead1=self.extra_overhead1, overhead2=self.extra_overhead2,
                lift1=self.extra_lifting1, lift2=self.extra_lifting2,
                trash1=self.extra_trash1, trash2=self.extra_trash2,
            )
            self.statusBar().showMessage(f"Экспорт: {path}")
            QMessageBox.information(self, "Экспорт", f"Смета экспортирована:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка экспорта", str(e))

    def _import_from_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", self.export_folder, "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, data_only=True)
            if "Смета" not in wb.sheetnames:
                QMessageBox.warning(self, "Ошибка", "Нет листа 'Смета'.")
                return
            ws = wb["Смета"]
            values = [list(row) for row in ws.iter_rows(values_only=True)]
            parsed = parse_exported_sheet(values)
            seq = parsed['sequence']
            if not seq:
                QMessageBox.warning(self, "Ошибка", "Не удалось распознать строки сметы.")
                return

            self.smeta_rows = []
            self.sections = []
            self.smeta_title = parsed['title'] if parsed['title'] else "Смета"
            self.title_edit.setText(self.smeta_title)

            self.extra_overhead1, self.extra_overhead2 = parsed['overhead']
            self.extra_lifting1, self.extra_lifting2 = parsed['lifting']
            self.extra_trash1, self.extra_trash2 = parsed['lifting_trash']

            for sp in [self.overhead1_spin, self.overhead2_spin,
                       self.lifting1_spin, self.lifting2_spin,
                       self.trash1_spin, self.trash2_spin]:
                sp.setValue(0)
            self.overhead1_spin.setValue(self.extra_overhead1)
            self.overhead2_spin.setValue(self.extra_overhead2)
            self.lifting1_spin.setValue(self.extra_lifting1)
            self.lifting2_spin.setValue(self.extra_lifting2)
            self.trash1_spin.setValue(self.extra_trash1)
            self.trash2_spin.setValue(self.extra_trash2)

            for item in seq:
                if item[0] == 'section':
                    self.sections.append(item[1])
                elif item[0] == 'work':
                    block = build_work_block(item[1], item[2], None, len(self.smeta_rows) + 1, self.db_mgr)
                    if block:
                        self.smeta_rows.extend(block)

            self.smeta_rows = rebuild_smeta(self.smeta_rows)
            self._update_table_from_rows()
            self._save_draft()

            self._refresh_works_table()
            self._refresh_materials_table()
            self._refresh_works_combo()
            self._refresh_link_works_combo()

            self.statusBar().showMessage(f"Импорт: {path}")
            QMessageBox.information(self, "Импорт", f"Смета импортирована:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка импорта", str(e))

    # =======================================================================
    # Слоты
    # =======================================================================
    def _on_title_changed(self, text):
        self.smeta_title = text.strip() if text else "Смета"

    def _on_extra_changed(self, _):
        self.extra_overhead1 = self.overhead1_spin.value()
        self.extra_overhead2 = self.overhead2_spin.value()
        self.extra_lifting1 = self.lifting1_spin.value()
        self.extra_lifting2 = self.lifting2_spin.value()
        self.extra_trash1 = self.trash1_spin.value()
        self.extra_trash2 = self.trash2_spin.value()
        self._save_draft()

    # =======================================================================
    # Справка / Настройки
    # =======================================================================
    def _show_about(self):
        QMessageBox.about(
            self, "О программе — Сметчик PRO 5.3",
            "Сметчик PRO 5.3\n\n"
            "Приложение для составления смет с поддержкой двух вариантов цены.\n\n"
            "Автор: pto_plus | Лицензия: MIT\n\n"
            "Технологии: Python 3, PyQt6, SQLite, pandas, openpyxl, xlsxwriter\n\n"
            "Функционал:\n"
            "• Справочник работ и материалов с ценами В1 и В2\n"
            "• Автоматический расчёт материалов по нормам расхода\n"
            "• Сравнение вариантов стоимости и расчёт экономии\n"
            "• Экспорт в Excel с формулами и форматированием\n"
            "• Импорт смет из Excel с восстановлением структуры\n"
            "• Черновик сметы с автосохранением\n"
            "• Управление разделами сметы\n"
            "• Дополнительные расходы: накладные, подъёмные механизмы, вывоз мусора\n"
            "• Создание и переключение между базами данных\n\n"
            "Горячие клавиши:\n"
            "Ctrl+S — Сохранить черновик\n"
            "Ctrl+E — Экспорт в Excel\n"
            "Ctrl+I — Импорт из Excel\n"
            "Ctrl+Z — Отменить действие\n"
            "Ctrl+Y — Повторить действие\n"
            "Ctrl+W — Добавить работу\n"
            "Ctrl+D — Добавить раздел\n"
            "Ctrl+Q — Выход\n\n"
            "Документация: https://docs.kodacode.ru\n"
            "Сообщество: https://t.me/kodacommunity",
        )

    def show_settings(self):
        dlg = SettingsDialog(self)
        dlg.folder_edit.setText(self.db_folder)
        dlg.export_edit.setText(self.export_folder)
        dlg.koeff_spin.setValue(self.koeff_price)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            folder, export, koeff = dlg.get_values()
            if folder:
                self.db_folder = folder
            if export:
                self.export_folder = export
            self.koeff_price = koeff
            self.statusBar().showMessage("Настройки сохранены")

    # =======================================================================
    # Закрытие
    # =======================================================================
    def closeEvent(self, event):
        self._save_draft()
        if self.db_mgr:
            self.db_mgr.flush()
            self.db_mgr.close()
        event.accept()


# ---------------------------------------------------------------------------
# Точка входа
# ---------------------------------------------------------------------------
def main():
    """Запускает приложение «Сметчик PRO 5.2»."""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = SmetaMainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
