"""
Модуль бизнес-логики для обработки смет.
Используется как в Telegram боте, так и в PWA приложении.
"""

import pandas as pd
import numpy as np
import re
import zipfile
import io
import os
from typing import Optional, Any, List, Union
from core.config_keywords import (
    WORK_KEYWORDS, MATERIAL_KEYWORDS, WORK_UNITS, PHYSICAL_UNITS,
    LABOR_UNITS, HEADER_PATTERNS, VERB_PATTERNS, MATERIAL_OVERRIDE, TECHNICAL_PATTERNS
)


KEYS = ["№ п/п", "Обоснование", "Наименование работ и затрат"]
COLS = ["№ п/п", "Обоснование", "Наименование работ и затрат", "Ед.изм.", "Кол-во на ед.", "Коэф.",
        "Кол-во всего", "Цена баз.", "Индекс", "Цена тек.", "Коэф.2", "Стоимость"]

norm = lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) \
                 else re.sub(r"\s+", " ", str(v)).strip()

def _numrow_map(row):
    d = {}
    for c, v in enumerate(row):
        s = norm(v)
        if re.fullmatch(r"\d+(\.0)?", s) and 1 <= int(float(s)) <= 12:
            d[int(float(s))] = c
    return d if len(d) >= 10 else None

def _header_map(row):
    cells = [norm(v) for v in row]
    p = [cells.index(k) for k in KEYS]
    u = next((c for c, s in enumerate(cells) if s.startswith("Единица")), p[2] + 1)
    return {1: p[0], 2: p[1], 3: p[2], 4: u, **{k: u + k - 4 for k in range(5, 13)}}

def parse_estimate(file_input: Union[str, io.BytesIO], **kwargs: Any) -> List[pd.DataFrame]:
    """
    Загружает и обрабатывает файл сметы Excel.
    
    Args:
        file_input: Путь к файлу сметы (str) или байтовый поток (BytesIO)
        **kwargs: Дополнительные параметры для парсинга
    
    Returns:
        Список DataFrame с данными сметы (по одному на каждый лист)
    """
    # Определяем тип файла и выбираем подходящий движок
    # .xlsx - новый формат (Office Open XML), требует openpyxl
    # .xls - старый бинарный формат, требует xlrd
    if isinstance(file_input, str):
        # Если передан путь, определяем по расширению
        if file_input.lower().endswith('.xls') and not file_input.lower().endswith('.xlsx'):
            xf = pd.ExcelFile(file_input, engine='xlrd')
        else:
            xf = pd.ExcelFile(file_input, engine='openpyxl')
    else:
        # Если передан BytesIO, пытаемся определить по первым байтам
        # Или можно передать engine явно через kwargs
        file_pos = file_input.tell()
        file_input.seek(0)
        header = file_input.read(8)
        file_input.seek(file_pos)
        
        # Проверяем сигнатуру файла:
        # .xls (BIFF8) начинается с D0 CF 11 E0 A1 B1 1A E1
        # .xlsx (Office Open XML) начинается с PK (50 4B 03 04)
        if header.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
            xf = pd.ExcelFile(file_input, engine='xlrd')
        elif header.startswith(b'PK\x03\x04'):
            xf = pd.ExcelFile(file_input, engine='openpyxl')
        else:
            # Пытаемся автоматически определить движок (pandas попробует угадать)
            xf = pd.ExcelFile(file_input)
    
    out = []
    for sh in xf.sheet_names:
        raw = xf.parse(sh, header=None).values
        rows, i = [], 0
        while i < len(raw):
            if not all(k in [norm(v) for v in raw[i]] for k in KEYS):
                i += 1
                continue
            cmap, start = None, i + 1
            for j in range(i + 1, min(i + 4, len(raw))):
                cmap = _numrow_map(raw[j])
                if cmap:
                    start = j + 1
                    break
            cmap = cmap or _header_map(raw[i])
            phys = [cmap[k] for k in range(1, 13)]
            i = start
            while i < len(raw):
                if all(k in [norm(v) for v in raw[i]] for k in KEYS):
                    break
                vals = [norm(raw[i][c]) if c < len(raw[i]) else "" for c in phys]
                if any(vals):
                    rows.append(vals)
                i += 1
        
        # Создаем DataFrame без жестких имен колонок - они будут определены позже
        if rows:
            df = pd.DataFrame(rows)
        else:
            df = pd.DataFrame()
        
        if not df.empty:
            df.attrs["sheet"] = sh
            out.append(df)
    return out

def extract_positions_by_sections(df: pd.DataFrame) -> pd.DataFrame:
    """
    Извлекает позиции сметы по разделам (позициям).
    
    Алгоритм:
    1. Находит строки где № п/п - целое число (1, 2, 3 без точек) - начало позиции
    2. Для каждой позиции ищет строку "Всего по позиции" в колонке "Наименование работ и затрат"
    3. Берет данные из первой строки (Наименование работ и затрат, Единица измерения, всего с учётом коэффициентов)
    4. Берет стоимость из последней строки ("Всего по позиции") из колонки "всего в текущем уровне цен"
    
    Args:
        df: DataFrame с данными сметы
    
    Returns:
        DataFrame с позициями: [№ п/п, Наименование работ и затрат, Единица измерения, всего с учётом коэффициентов, всего в текущем уровне цен]
    """
    if df.empty:
        return pd.DataFrame(columns=["№ п/п", "Наименование работ и затрат", "Единица измерения", "всего с учётом коэффициентов", "всего в текущем уровне цен"])
    
    result_rows = []
    current_position = None
    current_row_idx = None
    
    # Нормализуем названия колонок для поиска
    col_num = None
    col_name = None
    col_unit = None
    col_qty_total = None
    col_cost = None
    
    # Ищем соответствия колонок по оригинальным названиям
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if "№" in col_lower or "п/п" in col_lower or col_lower == "№":
            col_num = col
        elif "наименование" in col_lower:
            col_name = col
        elif "ед." in col_lower or "единиц" in col_lower or "изм" in col_lower:
            col_unit = col
        elif "всего" in col_lower and ("количеств" in col_lower or "кол-" in col_lower or "учётом" in col_lower):
            col_qty_total = col
        elif "всего в текущем" in col_lower or ("стоимость" in col_lower and "всего" in col_lower):
            col_cost = col
    
    # Если не нашли точных совпадений, используем стандартные имена
    if col_num is None:
        col_num = "№ п/п" if "№ п/п" in df.columns else df.columns[0]
    if col_name is None:
        col_name = "Наименование работ и затрат" if "Наименование работ и затрат" in df.columns else df.columns[2]
    if col_unit is None:
        col_unit = "Единица измерения" if "Единица измерения" in df.columns else df.columns[3]
    if col_qty_total is None:
        col_qty_total = "всего с учётом коэффициентов" if "всего с учётом коэффициентов" in df.columns else df.columns[6]
    if col_cost is None:
        col_cost = "всего в текущем уровне цен" if "всего в текущем уровне цен" in df.columns else df.columns[-1]
    
    def parse_cost(value):
        """Очистка стоимости от пробелов, символов валюты и преобразование в float"""
        if pd.isna(value) or value is None:
            return 0.0
        s = str(value).strip()
        # Удаляем символы валюты и пробелы
        s = re.sub(r'[₽\s]', '', s)
        s = re.sub(r'руб\.?', '', s, flags=re.IGNORECASE)
        s = s.replace(',', '.').strip()
        try:
            return float(s)
        except (ValueError, TypeError):
            return 0.0
    
    def is_integer_number(val):
        """Проверяет, является ли значение целым числом (без точек)"""
        if pd.isna(val) or val is None:
            return False
        s = str(val).strip()
        if not s:
            return False
        # Проверяем что строка состоит только из цифр
        return bool(re.fullmatch(r'\d+', s))
    
    for idx, row in df.iterrows():
        num_val = row.get(col_num, "")
        name_val = row.get(col_name, "")
        
        # Проверяем начало новой позиции (целое число в №)
        if is_integer_number(num_val):
            # Сохраняем текущую позицию если она была
            current_position = {
                '№ п/п': str(num_val).strip(),
                'Наименование работ и затрат': str(row.get(col_name, "")).strip(),
                'Единица измерения': str(row.get(col_unit, "")).strip(),
                'всего с учётом коэффициентов': str(row.get(col_qty_total, "")).strip(),
                'всего в текущем уровне цен': 0.0
            }
            current_row_idx = idx
        
        # Проверяем строку "Всего по позиции"
        elif current_position is not None and name_val is not None:
            name_str = str(name_val).lower().strip()
            if "всего по позиции" in name_str or "всего по разделу" in name_str:
                # Нашли итоговую строку позиции - берем стоимость
                cost_val = parse_cost(row.get(col_cost, 0))
                current_position['всего в текущем уровне цен'] = cost_val
                result_rows.append(current_position)
                current_position = None
                current_row_idx = None
    
    # Создаем результирующий DataFrame с новыми названиями колонок
    result_df = pd.DataFrame(result_rows, columns=["№ п/п", "Наименование работ и затрат", "Единица измерения", "всего с учётом коэффициентов", "всего в текущем уровне цен"])
    
    # Переименовываем колонки в итоговом виде
    rename_map = {
        "№ п/п": "№ п/п",
        "Наименование работ и затрат": "наименование",
        "Единица измерения": "ед. измерения",
        "всего с учётом коэффициентов": "объем",
        "всего в текущем уровне цен": "стоимость"
    }
    result_df = result_df.rename(columns=rename_map)
    return result_df


def calculate_totals_by_positions(df: pd.DataFrame) -> dict:
    """
    Вычисляет суммы по позициям сметы (без разделения на материалы/работы).
    
    Args:
        df: DataFrame с данными сметы
    
    Returns:
        Словарь с суммами: {'positions_total': float, 'positions_count': int}
    """
    
    if df.empty:
        return {
            'positions_total': 0.0,
            'positions_count': 0
        }
    
    # Преобразуем колонку "стоимость" в числовой формат
    positions_total = pd.to_numeric(df['стоимость'], errors='coerce').fillna(0).sum()
    
    return {
        'positions_total': positions_total,
        'positions_count': len(df)
    }


def export_estimate_to_excel(dfs, base_filename: str) -> tuple[bytes, str]:
    """
    Экспорт сметы в Excel файл с позициями.
    Если несколько листов в исходной смете, они объединяются в один Excel файл.
    
    Args:
        dfs: Список кортежей (имя_листа, DataFrame) с данными сметы
        base_filename: Базовое имя для файлов
    
    Returns:
        Кортеж из (байты Excel файла, имя файла)
    """
    excel_buffer = io.BytesIO()
    
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        for i, item in enumerate(dfs):
            # Поддерживаем как список кортежей (имя, df), так и список DataFrame
            if isinstance(item, tuple) and len(item) == 2:
                original_sheet_name, df = item
            else:
                df = item
                original_sheet_name = getattr(df, 'attrs', {}).get("sheet", f"sheet_{i+1}")
            
            safe_sheet_name = re.sub(r'[<>:"/\\|?*]', '', str(original_sheet_name))
            
            # Формируем имена листов с префиксом оригинального листа если их несколько
            prefix = f"{safe_sheet_name}_" if len(dfs) > 1 else ""
            
            # Лист: Позиции
            positions_sheet_name = f"{prefix}Позиции"
            if not df.empty:
                df.to_excel(writer, sheet_name=positions_sheet_name[:31], index=False)
            else:
                # Создаем пустой лист если нет позиций
                pd.DataFrame({'Наименование работ и затрат': ['Нет данных']}).to_excel(
                    writer, sheet_name=positions_sheet_name[:31], index=False
                )
    
    excel_content = excel_buffer.getvalue()
    excel_filename = f"{base_filename}_positions.xlsx"
    
    return excel_content, excel_filename
