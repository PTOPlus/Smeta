# -*- coding: utf-8 -*-
"""
db_manager.py
Менеджер базы данных с кэшированием в памяти и нормализованной структурой.
Использует Parquet для быстрого чтения/записи.
"""
import os
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq
from typing import Optional, Dict, List, Tuple
import time

# Нормализованная структура БД
WORKS_COLS = ['id', 'name', 'unit', 'price_1', 'price_2']
MATERIALS_COLS = ['id', 'name', 'unit', 'price_1', 'price_2']
WORK_MATERIALS_COLS = ['work_id', 'material_id', 'consumption_1', 'consumption_2']

# Старый формат (для миграции и совместимости)
LEGACY_COLS = ['Работа', 'Ед_изм_раб', 'Материал', 'Ед_изм', 
               'Расход_1', 'Цена_мат_1', 'Цена_раб_1',
               'Расход_2', 'Цена_мат_2', 'Цена_раб_2']

class DatabaseManager:
    """Менеджер БД с кэшированием и нормализованной структурой."""
    
    def __init__(self, db_folder: str, db_filename: str = 'smeta_db'):
        self.db_folder = db_folder
        self.db_filename = db_filename
        
        # Дефолтные пути (будут перезаписаны при автообнаружении)
        # Убираем расширение .xlsx если оно есть
        clean_name = db_filename.replace('.xlsx', '').replace('.parquet', '')
        
        self.works_path = os.path.join(db_folder, f'{clean_name}_works.parquet')
        self.materials_path = os.path.join(db_folder, f'{clean_name}_materials.parquet')
        self.work_materials_path = os.path.join(db_folder, f'{clean_name}_work_materials.parquet')
        self.legacy_path = os.path.join(db_folder, f'{clean_name}.xlsx')
        
        # Кэш в памяти
        self.works_cache: Optional[pd.DataFrame] = None
        self.materials_cache: Optional[pd.DataFrame] = None
        self.work_materials_cache: Optional[pd.DataFrame] = None
        
        # Флаги изменений (dirty flags)
        self.works_dirty = False
        self.materials_dirty = False
        self.work_materials_dirty = False
        
        # Загружаем или создаем БД
        self._init_db()
    
    def _init_db(self):
        """Инициализирует БД: ищет существующие Parquet или мигрирует с Excel."""
        # 1. Ищем любые Parquet-файлы в папке (автообнаружение)
        existing_parquet = self._find_existing_parquet()
        if existing_parquet:
            self.works_path = existing_parquet['works']
            self.materials_path = existing_parquet['materials']
            self.work_materials_path = existing_parquet['work_materials']
            self._load_from_parquet()
            return
        
        # 2. Ищем Excel-файл для миграции
        if os.path.exists(self.legacy_path):
            self._migrate_from_excel()
            return
        
        # 3. Создаём пустую БД
        self._create_empty_db()

    def _find_existing_parquet(self) -> Optional[Dict[str, str]]:
        """Ищет существующие Parquet-файлы в папке базы данных."""
        if not os.path.exists(self.db_folder):
            return None
        
        files = os.listdir(self.db_folder)
        
        # Ищем файлы по паттерну *_works.parquet, *_materials.parquet, *_work_materials.parquet
        works_files = [f for f in files if f.endswith('_works.parquet')]
        materials_files = [f for f in files if f.endswith('_materials.parquet')]
        work_materials_files = [f for f in files if f.endswith('_work_materials.parquet')]
        
        if works_files and materials_files and work_materials_files:
            # Берём первый найденный комплект (по алфавиту)
            base_name = works_files[0].replace('_works.parquet', '')
            return {
                'works': os.path.join(self.db_folder, f'{base_name}_works.parquet'),
                'materials': os.path.join(self.db_folder, f'{base_name}_materials.parquet'),
                'work_materials': os.path.join(self.db_folder, f'{base_name}_work_materials.parquet'),
            }
        
        return None
    
    def _load_from_parquet(self):
        """Загружает данные из Parquet файлов в кэш."""
        self.works_cache = pd.read_parquet(self.works_path)
        self.materials_cache = pd.read_parquet(self.materials_path)
        self.work_materials_cache = pd.read_parquet(self.work_materials_path)
        
        # Убедимся, что типы данных корректны
        for col in ['price_1', 'price_2', 'consumption_1', 'consumption_2']:
            if col in self.work_materials_cache.columns:
                self.work_materials_cache[col] = pd.to_numeric(self.work_materials_cache[col], errors='coerce').fillna(0.0)
        
        self.works_dirty = False
        self.materials_dirty = False
        self.work_materials_dirty = False
    
    def _migrate_from_excel(self):
        """Мигрирует данные с Excel на нормализованную структуру Parquet.
        O(n) — собираем списки словарей, один DataFrame в конце."""
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

        # Один вызов DataFrame на таблицу — O(n)
        self.works_cache = pd.DataFrame(works_list, columns=WORKS_COLS)
        self.materials_cache = pd.DataFrame(materials_list, columns=MATERIALS_COLS)
        self.work_materials_cache = pd.DataFrame(links_list, columns=WORK_MATERIALS_COLS)

        self._save_to_parquet()

        # Резервная копия исходного Excel
        backup_path = self.legacy_path.rsplit('.', 1)[0] + '_backup.xlsx'
        if not os.path.exists(backup_path):
            try:
                df.to_excel(backup_path, index=False)
                print(f"Миграция завершена. Резервная копия: {backup_path}")
            except Exception as e:
                print(f"Не удалось создать резервную копию Excel: {e}")
    
    def _create_empty_db(self):
        """Создает пустую нормализованную БД."""
        self.works_cache = pd.DataFrame(columns=WORKS_COLS)
        self.materials_cache = pd.DataFrame(columns=MATERIALS_COLS)
        self.work_materials_cache = pd.DataFrame(columns=WORK_MATERIALS_COLS)
        self._save_to_parquet()

    def _atomic_write_parquet(self, table, path: str, backup: bool = True):
        """Записывает Parquet-таблицу атомарно: tmp → os.replace.
        Перед перезаписью делает резервную копию .bak (если backup=True)."""
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp_path = path + '.tmp'

        # Резервная копия существующего файла
        if backup and os.path.exists(path):
            bak_path = path + '.bak'
            try:
                if os.path.exists(bak_path):
                    os.remove(bak_path)
                os.replace(path, bak_path)  # атомарный rename
            except OSError:
                pass  # не критично — продолжаем запись

        # Пишем во временный файл, затем атомарно заменяем целевой
        pq.write_table(table, tmp_path)
        os.replace(tmp_path, path)
    
    def _save_to_parquet(self):
        """Сохраняет кэш в Parquet файлы атомарно (tmp + os.replace)."""
        os.makedirs(self.db_folder, exist_ok=True)

        self._atomic_write_parquet(
            pa.Table.from_pandas(self.works_cache),
            self.works_path,
        )
        self._atomic_write_parquet(
            pa.Table.from_pandas(self.materials_cache),
            self.materials_path,
        )
        self._atomic_write_parquet(
            pa.Table.from_pandas(self.work_materials_cache),
            self.work_materials_path,
        )

        self.works_dirty = False
        self.materials_dirty = False
        self.work_materials_dirty = False
    
    def flush(self):
        """Принудительно сохраняет все изменения на диск."""
        if self.works_dirty or self.materials_dirty or self.work_materials_dirty:
            self._save_to_parquet()
    
    # --- Методы для работы с работами ---
    
    def get_works(self) -> pd.DataFrame:
        return self.works_cache.copy()
    
    def get_work_by_name(self, name: str) -> Optional[pd.Series]:
        mask = self.works_cache['name'].str.strip() == name.strip()
        result = self.works_cache[mask]
        return result.iloc[0] if not result.empty else None
    
    def add_work(self, name: str, unit: str, price_1: float, price_2: float) -> int:
        new_id = int(self.works_cache['id'].max()) + 1 if not self.works_cache.empty else 1
        new_row = pd.DataFrame([{
            'id': new_id, 'name': name, 'unit': unit,
            'price_1': price_1, 'price_2': price_2
        }])
        self.works_cache = pd.concat([self.works_cache, new_row], ignore_index=True)
        self.works_dirty = True
        return new_id
    
    def update_work(self, work_id: int, **kwargs):
        mask = self.works_cache['id'] == work_id
        for key, value in kwargs.items():
            if key in WORKS_COLS:
                self.works_cache.loc[mask, key] = value
        self.works_dirty = True
    
    def delete_work(self, work_id: int):
        self.works_cache = self.works_cache[self.works_cache['id'] != work_id]
        self.work_materials_cache = self.work_materials_cache[self.work_materials_cache['work_id'] != work_id]
        self.works_dirty = True
        self.work_materials_dirty = True
    
    # --- Методы для работы с материалами ---
    
    def get_materials(self) -> pd.DataFrame:
        return self.materials_cache.copy()
    
    def get_material_by_name(self, name: str) -> Optional[pd.Series]:
        mask = self.materials_cache['name'].str.strip() == name.strip()
        result = self.materials_cache[mask]
        return result.iloc[0] if not result.empty else None
    
    def add_material(self, name: str, unit: str, price_1: float, price_2: float) -> int:
        new_id = int(self.materials_cache['id'].max()) + 1 if not self.materials_cache.empty else 1
        new_row = pd.DataFrame([{
            'id': new_id, 'name': name, 'unit': unit,
            'price_1': price_1, 'price_2': price_2
        }])
        self.materials_cache = pd.concat([self.materials_cache, new_row], ignore_index=True)
        self.materials_dirty = True
        return new_id
    
    def update_material(self, material_id: int, **kwargs):
        mask = self.materials_cache['id'] == material_id
        for key, value in kwargs.items():
            if key in MATERIALS_COLS:
                self.materials_cache.loc[mask, key] = value
        self.materials_dirty = True
    
    def delete_material(self, material_id: int):
        self.materials_cache = self.materials_cache[self.materials_cache['id'] != material_id]
        self.work_materials_cache = self.work_materials_cache[self.work_materials_cache['material_id'] != material_id]
        self.materials_dirty = True
        self.work_materials_dirty = True
    
    # --- Методы для работы со связями ---
    
    def get_work_with_materials(self, work_name: str) -> Optional[Dict]:
        """Возвращает работу со всеми материалами (JOIN)."""
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
        mask = (self.work_materials_cache['work_id'] == work_id) & (self.work_materials_cache['material_id'] == material_id)
        self.work_materials_cache = self.work_materials_cache[~mask]
        self.work_materials_dirty = True
    
    # --- Методы для обратной совместимости ---
    
    def get_legacy_dataframe(self) -> pd.DataFrame:
        """Возвращает данные в старом формате (для совместимости с smeta_core.py)."""
        if self.work_materials_cache.empty:
            return pd.DataFrame(columns=LEGACY_COLS)
        
        result = []
        for _, link in self.work_materials_cache.iterrows():
            work = self.works_cache[self.works_cache['id'] == link['work_id']].iloc[0]
            mat = self.materials_cache[self.materials_cache['id'] == link['material_id']].iloc[0]
            
            result.append({
                'Работа': work['name'], 'Ед_изм_раб': work['unit'],
                'Материал': mat['name'], 'Ед_изм': mat['unit'],
                'Расход_1': link['consumption_1'], 'Цена_мат_1': mat['price_1'], 'Цена_раб_1': work['price_1'],
                'Расход_2': link['consumption_2'], 'Цена_мат_2': mat['price_2'], 'Цена_раб_2': work['price_2']
            })
        
        return pd.DataFrame(result, columns=LEGACY_COLS)
    
    def save_legacy_dataframe(self, df: pd.DataFrame):
        """Сохраняет данные в старом формате (мигрирует на нормализованную структуру).
        O(n) — списки словарей, один DataFrame в конце."""
        works_list = []
        materials_list = []
        links_list = []

        work_id = 1
        mat_id = 1
        work_id_map = {}
        mat_id_map = {}

        for _, row in df.iterrows():
            work_name = str(row['Работа']).strip()
            mat_name = str(row['Материал']).strip()

            if work_name not in work_id_map:
                work_id_map[work_name] = work_id
                works_list.append({
                    'id': work_id,
                    'name': work_name,
                    'unit': str(row['Ед_изм_раб']).strip(),
                    'price_1': float(row['Цена_раб_1']) if pd.notna(row['Цена_раб_1']) else 0.0,
                    'price_2': float(row['Цена_раб_2']) if pd.notna(row['Цена_раб_2']) else 0.0,
                })
                work_id += 1

            if mat_name not in ('', '-', '0') and mat_name not in mat_id_map:
                mat_id_map[mat_name] = mat_id
                materials_list.append({
                    'id': mat_id,
                    'name': mat_name,
                    'unit': str(row['Ед_изм']).strip(),
                    'price_1': float(row['Цена_мат_1']) if pd.notna(row['Цена_мат_1']) else 0.0,
                    'price_2': float(row['Цена_мат_2']) if pd.notna(row['Цена_мат_2']) else 0.0,
                })
                mat_id += 1

            if mat_name not in ('', '-', '0'):
                links_list.append({
                    'work_id': work_id_map[work_name],
                    'material_id': mat_id_map[mat_name],
                    'consumption_1': float(row['Расход_1']) if pd.notna(row['Расход_1']) else 0.0,
                    'consumption_2': float(row['Расход_2']) if pd.notna(row['Расход_2']) else 0.0,
                })

        self.works_cache = pd.DataFrame(works_list, columns=WORKS_COLS)
        self.materials_cache = pd.DataFrame(materials_list, columns=MATERIALS_COLS)
        self.work_materials_cache = pd.DataFrame(links_list, columns=WORK_MATERIALS_COLS)

        self.works_dirty = True
        self.materials_dirty = True
        self.work_materials_dirty = True
        self._save_to_parquet()
    
    # --- Поиск и фильтрация ---
    
    def search_works(self, query: str) -> pd.DataFrame:
        mask = self.works_cache['name'].str.lower().str.contains(query.lower(), na=False)
        return self.works_cache[mask]
    
    def search_materials(self, query: str) -> pd.DataFrame:
        mask = self.materials_cache['name'].str.lower().str.contains(query.lower(), na=False)
        return self.materials_cache[mask]
    
    def get_works_by_material(self, material_name: str) -> pd.DataFrame:
        mat = self.get_material_by_name(material_name)
        if mat is None: return pd.DataFrame()
        
        mat_id = mat['id']
        mask = self.work_materials_cache['material_id'] == mat_id
        work_ids = self.work_materials_cache[mask]['work_id'].unique()
        return self.works_cache[self.works_cache['id'].isin(work_ids)]