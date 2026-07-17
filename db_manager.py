# -*- coding: utf-8 -*-
"""
db_manager.py
Менеджер базы данных с кэшированием в памяти и нормализованной структурой.
Использует SQLite для хранения данных (быстрое чтение/запись, поддержка связей).
"""
import os
import sqlite3
import pandas as pd
from typing import Optional, Dict

# Нормализованная структура БД (колонки для DataFrame-представления)
WORKS_COLS = ['id', 'name', 'unit', 'price_1', 'price_2']
MATERIALS_COLS = ['id', 'name', 'unit', 'price_1', 'price_2']
WORK_MATERIALS_COLS = ['work_id', 'material_id', 'consumption_1', 'consumption_2']

# Старый формат (для миграции и совместимости)
LEGACY_COLS = ['Работа', 'Ед_изм_раб', 'Материал', 'Ед_изм',
               'Расход_1', 'Цена_мат_1', 'Цена_раб_1',
               'Расход_2', 'Цена_мат_2', 'Цена_раб_2']


class DatabaseManager:
    """Менеджер базы данных на основе SQLite с кэшированием в памяти.
    
    Обеспечивает хранение работ, материалов и связей между ними в нормализованной
    структуре. Данные кэшируются в pandas DataFrame для быстрого доступа,
    а при изменении флагов dirty автоматически сохраняются в SQLite.
    
    Поддерживает миграцию из legacy-формата Excel (.xlsx) в SQLite.
    
    Attributes:
        db_folder (str): путь к папке с базой данных.
        db_filename (str): имя файла базы (без расширения).
        sqlite_path (str): полный путь к SQLite-файлу.
        legacy_path (str): полный путь к legacy Excel-файлу.
        works_cache (pd.DataFrame): кэш таблицы работ.
        materials_cache (pd.DataFrame): кэш таблицы материалов.
        work_materials_cache (pd.DataFrame): кэш таблицы связей.
        works_dirty (bool): флаг изменений в кэше работ.
        materials_dirty (bool): флаг изменений в кэше материалов.
        work_materials_dirty (bool): флаг изменений в кэше связей.
        conn (sqlite3.Connection): соединение с SQLite.
    """

    def __init__(self, db_folder: str, db_filename: str = 'smeta_db'):
        """Инициализирует менеджер базы данных.
        
        Args:
            db_folder (str): путь к папке, где хранятся файлы базы данных.
            db_filename (str): имя базы данных (без расширения). По умолчанию 'smeta_db'.
        """
        self.db_folder = db_folder
        self.db_filename = db_filename

        # Убираем расширение если оно есть
        clean_name = db_filename.replace('.xlsx', '').replace('.parquet', '').replace('.db', '')

        self.sqlite_path = os.path.join(db_folder, f'{clean_name}.db')
        self.legacy_path = os.path.join(db_folder, f'{clean_name}.xlsx')

        # Кэш в памяти
        self.works_cache: Optional[pd.DataFrame] = None
        self.materials_cache: Optional[pd.DataFrame] = None
        self.work_materials_cache: Optional[pd.DataFrame] = None

        # Флаги изменений (dirty flags)
        self.works_dirty = False
        self.materials_dirty = False
        self.work_materials_dirty = False

        # Подключаемся к SQLite
        self.conn = sqlite3.connect(self.sqlite_path)
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA foreign_keys=ON")

        # Загружаем или создаём БД
        self._init_db()

    def _init_db(self):
        """Инициализирует базу данных: ищет существующую SQLite или мигрирует с Excel.
        
        Алгоритм:
            1. Если SQLite-файл существует — загружает данные из него.
            2. Иначе ищет legacy Excel-файл и выполняет миграцию.
            3. Если ничего не найдено — создаёт пустую базу данных.
        """
        # 1. Если SQLite уже существует — просто загружаем данные
        if os.path.exists(self.sqlite_path):
            self._load_from_sqlite()
            return

        # 2. Ищем Excel-файл для миграции
        if os.path.exists(self.legacy_path):
            self._migrate_from_excel()
            return

        # 3. Создаём пустую БД
        self._create_empty_db()

    def _create_tables(self):
        """Создаёт таблицы SQLite (works, materials, work_materials), если они не существуют.
        
        Таблицы:
            - works: работы с id, name, unit, price_1, price_2
            - materials: материалы с id, name, unit, price_1, price_2
            - work_materials: связи между работами и материалами с FK и CASCADE
        """
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS works (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                unit TEXT,
                price_1 REAL DEFAULT 0.0,
                price_2 REAL DEFAULT 0.0
            );

            CREATE TABLE IF NOT EXISTS materials (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                unit TEXT,
                price_1 REAL DEFAULT 0.0,
                price_2 REAL DEFAULT 0.0
            );

            CREATE TABLE IF NOT EXISTS work_materials (
                work_id INTEGER NOT NULL,
                material_id INTEGER NOT NULL,
                consumption_1 REAL DEFAULT 0.0,
                consumption_2 REAL DEFAULT 0.0,
                PRIMARY KEY (work_id, material_id),
                FOREIGN KEY (work_id) REFERENCES works(id) ON DELETE CASCADE,
                FOREIGN KEY (material_id) REFERENCES materials(id) ON DELETE CASCADE
            );
        """)

    def _load_from_sqlite(self):
        """Загружает данные из SQLite в кэш в памяти.
        
        Создаёт таблицы (если не существуют), читает все три таблицы в DataFrame
        и приводит числовые колонки к корректным типам.
        """
        self._create_tables()

        # Читаем таблицы
        self.works_cache = pd.read_sql_query("SELECT * FROM works", self.conn)
        self.materials_cache = pd.read_sql_query("SELECT * FROM materials", self.conn)
        self.work_materials_cache = pd.read_sql_query("SELECT * FROM work_materials", self.conn)

        # Убедимся, что типы данных корректны
        for col in ['price_1', 'price_2', 'consumption_1', 'consumption_2']:
            if col in self.works_cache.columns:
                self.works_cache[col] = pd.to_numeric(self.works_cache[col], errors='coerce').fillna(0.0)
            if col in self.materials_cache.columns:
                self.materials_cache[col] = pd.to_numeric(self.materials_cache[col], errors='coerce').fillna(0.0)
            if col in self.work_materials_cache.columns:
                self.work_materials_cache[col] = pd.to_numeric(self.work_materials_cache[col], errors='coerce').fillna(0.0)

        self.works_dirty = False
        self.materials_dirty = False
        self.work_materials_dirty = False

    def _migrate_from_excel(self):
        """Мигрирует данные из legacy Excel-файла в нормализованную структуру SQLite.
        
        Преобразует плоскую таблицу Excel в три нормализованные таблицы:
        works, materials, work_materials. Создаёт резервную копию исходного Excel.
        
        Args:
            legacy_path: путь к Excel-файлу (используется self.legacy_path).
            
        Returns:
            None: данные сохраняются в кэш и записываются в SQLite.
        """
        print(f"Миграция с Excel: {self.legacy_path}")

        try:
            df = pd.read_excel(self.legacy_path)
        except Exception as e:
            print(f"Ошибка чтения Excel: {e}")
            self._create_empty_db()
            return

        required_cols = ['Работа', 'Ед_изм_раб', 'Цена_раб_1', 'Цена_раб_2']
        if not all(col in df.columns for col in required_cols):
            print("Неверный формат Excel, создаем пустую БД")
            self._create_empty_db()
            return

        works_list = []
        materials_list = []
        links_list = []

        work_id = 1
        mat_id = 1
        work_id_map = {}
        mat_id_map = {}

        for _, row in df.iterrows():
            work_name = str(row['Работа']).strip()
            if not work_name:
                continue
            mat_name = str(row.get('Материал', '-')).strip()

            # Работа — добавляем один раз
            if work_name not in work_id_map:
                work_id_map[work_name] = work_id
                works_list.append({
                    'id': work_id,
                    'name': work_name,
                    'unit': str(row.get('Ед_изм_раб', '')).strip(),
                    'price_1': float(row['Цена_раб_1']) if pd.notna(row.get('Цена_раб_1')) else 0.0,
                    'price_2': float(row['Цена_раб_2']) if pd.notna(row.get('Цена_раб_2')) else 0.0,
                })
                work_id += 1

            # Материал + связь
            if mat_name not in ('', '-', '0', 'nan'):
                if mat_name not in mat_id_map:
                    mat_id_map[mat_name] = mat_id
                    materials_list.append({
                        'id': mat_id,
                        'name': mat_name,
                        'unit': str(row.get('Ед_изм', '')).strip(),
                        'price_1': float(row.get('Цена_мат_1', 0)) if pd.notna(row.get('Цена_мат_1')) else 0.0,
                        'price_2': float(row.get('Цена_мат_2', 0)) if pd.notna(row.get('Цена_мат_2')) else 0.0,
                    })
                    mat_id += 1

                links_list.append({
                    'work_id': work_id_map[work_name],
                    'material_id': mat_id_map[mat_name],
                    'consumption_1': float(row.get('Расход_1', 0)) if pd.notna(row.get('Расход_1')) else 0.0,
                    'consumption_2': float(row.get('Расход_2', 0)) if pd.notna(row.get('Расход_2')) else 0.0,
                })

        # Создаём DataFrame из собранных данных
        self.works_cache = pd.DataFrame(works_list, columns=WORKS_COLS) if works_list else pd.DataFrame(columns=WORKS_COLS)
        self.materials_cache = pd.DataFrame(materials_list, columns=MATERIALS_COLS) if materials_list else pd.DataFrame(columns=MATERIALS_COLS)
        self.work_materials_cache = pd.DataFrame(links_list, columns=WORK_MATERIALS_COLS) if links_list else pd.DataFrame(columns=WORK_MATERIALS_COLS)

        # Сохраняем в SQLite
        self._save_to_sqlite()

        # Резервная копия исходного Excel
        backup_path = self.legacy_path.rsplit('.', 1)[0] + '_backup.xlsx'
        if not os.path.exists(backup_path):
            try:
                df.to_excel(backup_path, index=False)
                print(f"Миграция завершена. Резервная копия: {backup_path}")
            except Exception as e:
                print(f"Не удалось создать резервную копию Excel: {e}")

    def _create_empty_db(self):
        """Создаёт пустую нормализованную базу данных.
        
        Инициализирует таблицы и создаёт пустые DataFrame для всех кэшей.
        """
        self._create_tables()
        self.works_cache = pd.DataFrame(columns=WORKS_COLS)
        self.materials_cache = pd.DataFrame(columns=MATERIALS_COLS)
        self.work_materials_cache = pd.DataFrame(columns=WORK_MATERIALS_COLS)
        self._save_to_sqlite()

    def _save_to_sqlite(self):
        """Сохраняет все кэшированные данные в SQLite.
        
        Перед записью создаёт резервную копию существующего .db файла.
        Отключает foreign keys на время перезаписи таблиц, затем включает обратно.
        Сбрасывает все dirty-флаги после успешного сохранения.
        """
        os.makedirs(self.db_folder, exist_ok=True)

        # Делаем резервную копию перед перезаписью
        if os.path.exists(self.sqlite_path):
            bak_path = self.sqlite_path + '.bak'
            try:
                if os.path.exists(bak_path):
                    os.remove(bak_path)
                os.replace(self.sqlite_path, bak_path)
            except OSError:
                pass

        self._create_tables()

        # ВАЖНО: отключаем foreign keys для корректной замены таблиц
        self.conn.execute("PRAGMA foreign_keys=OFF")

        try:
            if not self.works_cache.empty:
                self.works_cache.to_sql('works', self.conn, if_exists='replace', index=False)
            else:
                self.conn.execute("DELETE FROM works")

            if not self.materials_cache.empty:
                self.materials_cache.to_sql('materials', self.conn, if_exists='replace', index=False)
            else:
                self.conn.execute("DELETE FROM materials")

            if not self.work_materials_cache.empty:
                self.work_materials_cache.to_sql('work_materials', self.conn, if_exists='replace', index=False)
            else:
                self.conn.execute("DELETE FROM work_materials")
        finally:
            self.conn.execute("PRAGMA foreign_keys=ON")

        self.conn.commit()

        self.works_dirty = False
        self.materials_dirty = False
        self.work_materials_dirty = False

    def flush(self):
        """Принудительно сохраняет все изменения на диск.
        
        Проверяет dirty-флаги и вызывает _save_to_sqlite() при наличии изменений.
        """
        if self.works_dirty or self.materials_dirty or self.work_materials_dirty:
            self._save_to_sqlite()

    # -------------------------------------------------------------------------
    # Методы для работы с работами
    # -------------------------------------------------------------------------

    def get_works(self) -> pd.DataFrame:
        """Возвращает копию кэша работ.
        
        Returns:
            pd.DataFrame: копия таблицы работ.
        """
        return self.works_cache.copy()

    def get_work_by_name(self, name: str) -> Optional[pd.Series]:
        """Находит работу в кэше по точному совпадению имени.
        
        Args:
            name (str): название работы для поиска.
            
        Returns:
            pd.Series или None: строка работы, если найдена, иначе None.
        """
        mask = self.works_cache['name'].str.strip() == name.strip()
        result = self.works_cache[mask]
        return result.iloc[0] if not result.empty else None

    def add_work(self, name: str, unit: str, price_1: float, price_2: float) -> int:
        """Добавляет новую работу в кэш и помечает его как изменённый.
        
        Args:
            name (str): название работы.
            unit (str): единица измерения.
            price_1 (float): цена варианта 1.
            price_2 (float): цена варианта 2.
            
        Returns:
            int: id вновь созданной работы.
        """
        new_id = int(self.works_cache['id'].max()) + 1 if not self.works_cache.empty else 1
        new_row = pd.DataFrame([{
            'id': new_id, 'name': name, 'unit': unit,
            'price_1': price_1, 'price_2': price_2
        }])
        self.works_cache = pd.concat([self.works_cache, new_row], ignore_index=True)
        self.works_dirty = True
        return new_id

    def update_work(self, work_id: int, **kwargs):
        """Обновляет поля работы по её id.
        
        Args:
            work_id (int): id работы для обновления.
            **kwargs: словари с полями для обновления (key, value).
        """
        mask = self.works_cache['id'] == work_id
        for key, value in kwargs.items():
            if key in WORKS_COLS:
                self.works_cache.loc[mask, key] = value
        self.works_dirty = True

    def delete_work(self, work_id: int):
        """Удаляет работу и все её связи из кэша.
        
        Args:
            work_id (int): id работы для удаления.
        """
        self.works_cache = self.works_cache[self.works_cache['id'] != work_id]
        self.work_materials_cache = self.work_materials_cache[self.work_materials_cache['work_id'] != work_id]
        self.works_dirty = True
        self.work_materials_dirty = True

    # -------------------------------------------------------------------------
    # Методы для работы с материалами
    # -------------------------------------------------------------------------

    def get_materials(self) -> pd.DataFrame:
        """Возвращает копию кэша материалов.
        
        Returns:
            pd.DataFrame: копия таблицы материалов.
        """
        return self.materials_cache.copy()

    def get_material_by_name(self, name: str) -> Optional[pd.Series]:
        """Находит материал в кэше по точному совпадению имени.
        
        Args:
            name (str): название материала для поиска.
            
        Returns:
            pd.Series или None: строка материала, если найдена, иначе None.
        """
        mask = self.materials_cache['name'].str.strip() == name.strip()
        result = self.materials_cache[mask]
        return result.iloc[0] if not result.empty else None

    def add_material(self, name: str, unit: str, price_1: float, price_2: float) -> int:
        """Добавляет новый материал в кэш и помечает его как изменённый.
        
        Args:
            name (str): название материала.
            unit (str): единица измерения.
            price_1 (float): цена варианта 1.
            price_2 (float): цена варианта 2.
            
        Returns:
            int: id вновь созданного материала.
        """
        new_id = int(self.materials_cache['id'].max()) + 1 if not self.materials_cache.empty else 1
        new_row = pd.DataFrame([{
            'id': new_id, 'name': name, 'unit': unit,
            'price_1': price_1, 'price_2': price_2
        }])
        self.materials_cache = pd.concat([self.materials_cache, new_row], ignore_index=True)
        self.materials_dirty = True
        return new_id

    def update_material(self, material_id: int, **kwargs):
        """Обновляет поля материала по его id.
        
        Args:
            material_id (int): id материала для обновления.
            **kwargs: словари с полями для обновления (key, value).
        """
        mask = self.materials_cache['id'] == material_id
        for key, value in kwargs.items():
            if key in MATERIALS_COLS:
                self.materials_cache.loc[mask, key] = value
        self.materials_dirty = True

    def delete_material(self, material_id: int):
        """Удаляет материал и все его связи из кэша.
        
        Args:
            material_id (int): id материала для удаления.
        """
        self.materials_cache = self.materials_cache[self.materials_cache['id'] != material_id]
        self.work_materials_cache = self.work_materials_cache[self.work_materials_cache['material_id'] != material_id]
        self.materials_dirty = True
        self.work_materials_dirty = True

    # -------------------------------------------------------------------------
    # Методы для работы со связями
    # -------------------------------------------------------------------------

    def get_work_with_materials(self, work_name: str) -> Optional[Dict]:
        """Возвращает работу со всеми привязанными материалами (JOIN).
        
        Args:
            work_name (str): название работы для поиска.
            
        Returns:
            dict или None: словарь с ключами 'work' (pd.Series) и 
                'materials' (list[dict]), или None если работа не найдена.
        """
        work = self.get_work_by_name(work_name)
        if work is None:
            return None

        work_id = work['id']
        mask = self.work_materials_cache['work_id'] == work_id
        links = self.work_materials_cache[mask]

        materials = []
        for _, link in links.iterrows():
            mat = self.materials_cache[self.materials_cache['id'] == link['material_id']].iloc[0]
            materials.append({
                'name': mat['name'], 'unit': mat['unit'],
                'price_1': mat['price_1'], 'price_2': mat['price_2'],
                'consumption_1': link['consumption_1'], 'consumption_2': link['consumption_2']
            })

        return {'work': work, 'materials': materials}

    def add_work_material_link(self, work_id: int, material_id: int, consumption_1: float, consumption_2: float):
        """Добавляет или обновляет связь между работой и материалом.
        
        Если связь уже существует — обновляет нормы расхода. Иначе создаёт новую.
        
        Args:
            work_id (int): id работы.
            material_id (int): id материала.
            consumption_1 (float): норма расхода варианта 1.
            consumption_2 (float): норма расхода варианта 2.
        """
        mask = (self.work_materials_cache['work_id'] == work_id) & (self.work_materials_cache['material_id'] == material_id)

        if mask.any():
            self.work_materials_cache.loc[mask, 'consumption_1'] = consumption_1
            self.work_materials_cache.loc[mask, 'consumption_2'] = consumption_2
        else:
            new_row = pd.DataFrame([{
                'work_id': work_id, 'material_id': material_id,
                'consumption_1': consumption_1, 'consumption_2': consumption_2
            }])
            self.work_materials_cache = pd.concat([self.work_materials_cache, new_row], ignore_index=True)

        self.work_materials_dirty = True

    def remove_work_material_link(self, work_id: int, material_id: int):
        """Удаляет связь между конкретной работой и материалом.
        
        Args:
            work_id (int): id работы.
            material_id (int): id материала.
        """
        mask = (self.work_materials_cache['work_id'] == work_id) & (self.work_materials_cache['material_id'] == material_id)
        self.work_materials_cache = self.work_materials_cache[~mask]
        self.work_materials_dirty = True

    def delete_work_material_links_by_work(self, work_id: int):
        """Удаляет ВСЕ связи материалов для указанной работы.
        
        Используется при редактировании работы — сначала удаляются старые
        связи, затем создаются новые.
        
        Args:
            work_id (int): id работы, связи которой нужно удалить.
        """
        self.work_materials_cache = self.work_materials_cache[self.work_materials_cache['work_id'] != work_id]
        self.work_materials_dirty = True

    # -------------------------------------------------------------------------
    # Методы для обратной совместимости
    # -------------------------------------------------------------------------

    def get_legacy_dataframe(self) -> pd.DataFrame:
        """Возвращает данные в старом плоском формате для совместимости.
        
        Преобразует нормализованные таблицы (works, materials, work_materials)
        в плоский DataFrame с LEGACY_COLS. Работы без материалов возвращаются
        с '-' в поле 'Материал'.
        
        Returns:
            pd.DataFrame: DataFrame в старом формате (LEGACY_COLS).
        """
        if self.works_cache.empty:
            return pd.DataFrame(columns=LEGACY_COLS)

        result = []
        for _, work in self.works_cache.iterrows():
            work_id = work['id']
            mask = self.work_materials_cache['work_id'] == work_id
            links = self.work_materials_cache[mask]

            if links.empty:
                # Работа без материалов — возвращаем строку с '-'
                result.append({
                    'Работа': work['name'], 'Ед_изм_раб': work['unit'],
                    'Материал': '-', 'Ед_изм': '-',
                    'Расход_1': 0.0, 'Цена_мат_1': 0.0, 'Цена_раб_1': work['price_1'],
                    'Расход_2': 0.0, 'Цена_мат_2': 0.0, 'Цена_раб_2': work['price_2']
                })
            else:
                for _, link in links.iterrows():
                    mat = self.materials_cache[self.materials_cache['id'] == link['material_id']].iloc[0]
                    result.append({
                        'Работа': work['name'], 'Ед_изм_раб': work['unit'],
                        'Материал': mat['name'], 'Ед_изм': mat['unit'],
                        'Расход_1': link['consumption_1'], 'Цена_мат_1': mat['price_1'], 'Цена_раб_1': work['price_1'],
                        'Расход_2': link['consumption_2'], 'Цена_мат_2': mat['price_2'], 'Цена_раб_2': work['price_2']
                    })

        return pd.DataFrame(result, columns=LEGACY_COLS)

    def save_legacy_dataframe(self, df: pd.DataFrame):
        """Добавляет или обновляет записи из legacy DataFrame в нормализованной БД.
        
        Проходит по строкам переданного DataFrame и для каждой строки:
        - находит или создаёт работу
        - находит или создаёт материал
        - создаёт или обновляет связь между ними
        
        Args:
            df (pd.DataFrame): DataFrame в старом плоском формате (LEGACY_COLS).
        """
        for _, row in df.iterrows():
            work_name = str(row['Работа']).strip()
            if not work_name:
                continue
            mat_name = str(row['Материал']).strip()

            unit_w = str(row['Ед_изм_раб']).strip()
            price_w1 = float(row['Цена_раб_1']) if pd.notna(row['Цена_раб_1']) else 0.0
            price_w2 = float(row['Цена_раб_2']) if pd.notna(row['Цена_раб_2']) else 0.0

            work = self.get_work_by_name(work_name)
            if work is None:
                work_id = self.add_work(work_name, unit_w, price_w1, price_w2)
            else:
                work_id = int(work['id'])
                self.update_work(work_id, unit=unit_w, price_1=price_w1, price_2=price_w2)

            if mat_name not in ('', '-', '0', 'nan'):
                unit_m = str(row['Ед_изм']).strip()
                price_m1 = float(row['Цена_мат_1']) if pd.notna(row['Цена_мат_1']) else 0.0
                price_m2 = float(row['Цена_мат_2']) if pd.notna(row['Цена_мат_2']) else 0.0

                mat = self.get_material_by_name(mat_name)
                if mat is None:
                    mat_id = self.add_material(mat_name, unit_m, price_m1, price_m2)
                else:
                    mat_id = int(mat['id'])
                    self.update_material(mat_id, unit=unit_m, price_1=price_m1, price_2=price_m2)

                cons1 = float(row['Расход_1']) if pd.notna(row['Расход_1']) else 0.0
                cons2 = float(row['Расход_2']) if pd.notna(row['Расход_2']) else 0.0
                self.add_work_material_link(work_id, mat_id, cons1, cons2)

        self._save_to_sqlite()

    # -------------------------------------------------------------------------
    # Поиск и фильтрация
    # -------------------------------------------------------------------------

    def search_works(self, query: str) -> pd.DataFrame:
        """Ищет работы по подстроке в названии (без учёта регистра).
        
        Args:
            query (str): поисковый запрос.
            
        Returns:
            pd.DataFrame: строки, содержащие запрос в названии.
        """
        mask = self.works_cache['name'].str.lower().str.contains(query.lower(), na=False)
        return self.works_cache[mask]

    def search_materials(self, query: str) -> pd.DataFrame:
        """Ищет материалы по подстроке в названии (без учёта регистра).
        
        Args:
            query (str): поисковый запрос.
            
        Returns:
            pd.DataFrame: строки, содержащие запрос в названии.
        """
        mask = self.materials_cache['name'].str.lower().str.contains(query.lower(), na=False)
        return self.materials_cache[mask]

    def get_works_by_material(self, material_name: str) -> pd.DataFrame:
        """Находит все работы, к которым привязан указанный материал.
        
        Args:
            material_name (str): название материала для поиска.
            
        Returns:
            pd.DataFrame: DataFrame с найденными работами. Пустой DataFrame, если материал не найден.
        """
        mat = self.get_material_by_name(material_name)
        if mat is None:
            return pd.DataFrame()

        mat_id = mat['id']
        mask = self.work_materials_cache['material_id'] == mat_id
        work_ids = self.work_materials_cache[mask]['work_id'].unique()
        return self.works_cache[self.works_cache['id'].isin(work_ids)]

    # -------------------------------------------------------------------------
    # Закрытие соединения
    # -------------------------------------------------------------------------

    def close(self):
        """Закрывает соединение с SQLite и сохраняет изменения.
        
        Вызывает flush() для сохранения всех изменённых данных перед закрытием.
        """
        self.flush()
        if self.conn:
            self.conn.close()

    def __del__(self):
        """Деструктор — гарантирует закрытие соединения с БД при удалении объекта."""
        try:
            self.close()
        except Exception:
            pass
