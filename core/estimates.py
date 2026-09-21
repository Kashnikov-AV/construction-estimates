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


def export_estimates_to_csv(dfs: List[pd.DataFrame], base_filename: str) -> bytes:
    """
    Экспорт списка DataFrame в ZIP-архив с CSV файлами.
    
    Args:
        dfs: Список DataFrame с данными сметы
        base_filename: Базовое имя для файлов (без расширения)
    
    Returns:
        Байты ZIP-архива с CSV файлами (или байты CSV если один лист)
    """
    # Если один лист, возвращаем просто CSV без архивации
    if len(dfs) == 1:
        df = dfs[0]
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False, sep=',', encoding='utf-8-sig')
        return csv_buffer.getvalue().encode('utf-8-sig')
    
    # Если несколько листов, упаковываем в ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for i, df in enumerate(dfs):
            sheet_name = df.attrs.get("sheet", f"sheet_{i+1}")
            # Очищаем имя листа от недопустимых символов для имени файла
            safe_sheet_name = re.sub(r'[<>:"/\\|?*]', '', sheet_name)
            # Формируем имя файла: {имя_excel}_{имя_листа}.csv (без транслитерации)
            csv_filename = f"{base_filename}_{safe_sheet_name}.csv"
            
            csv_buffer = io.StringIO()
            df.to_csv(csv_buffer, index=False, sep=',', encoding='utf-8-sig')
            csv_content = csv_buffer.getvalue().encode('utf-8-sig')
            
            zip_file.writestr(csv_filename, csv_content)
    
    return zip_buffer.getvalue()




def classify_estimate_items(df: pd.DataFrame) -> pd.DataFrame:
    """
    Классифицирует позиции сметы на работы/услуги и материалы.
    
    Алгоритм классификации (приоритетный порядок):
    1. Единицы измерения труда (чел.-ч, маш.-ч) — работа
    2. Укрупнённые единицы (100 м2, 10 шт и т.п.) — работа
    3. Проценты (%) — накладные расходы/сметная прибыль (не материал)
    4. Проверка на заголовки/итоги по HEADER_PATTERNS
    5. Технические расчеты и объемы (TECHNICAL_PATTERNS) — Заголовок/Итог
    6. Ключевые слова-процессы в наименовании из WORK_KEYWORDS — работа
    7. Ключевые слова-материалы из MATERIAL_KEYWORDS — материал
    8. Глагольные формы из VERB_PATTERNS — работа
    9. MATERIAL_OVERRIDE — приоритет материала даже при наличии глаголов
    10. Физические единицы без маркеров — предположительно материал
    11. Всё остальное — требует уточнения
    
    Args:
        df: DataFrame с данными сметы (колонки: "Наименование", "Ед.изм.")
    
    Returns:
        DataFrame с добавленной колонкой "Категория" со значениями:
        - "Работа" — работы и услуги
        - "Материал" — материальные ресурсы
        - "НР/СП" — накладные расходы и сметная прибыль
        - "Заголовок/Итог" — служебные строки и технические расчеты
        - "Требует уточнения" — сложные случаи
    """
    result_df = df.copy()
    categories = []
    
    for idx, row in result_df.iterrows():
        name = str(row.get("Наименование", "")).lower()
        unit = str(row.get("Ед.изм.", "")).strip()
        
        # Пустые или служебные строки
        if not name or name in ['', ' ', '-', 'итого', 'всего', 'ндс']:
            categories.append("Заголовок/Итог")
            continue
        
        # 1. Единицы измерения труда — работа
        if unit in LABOR_UNITS:
            categories.append("Работа")
            continue
        
        # 2. Проверка ключевых слов-материалов из MATERIAL_KEYWORDS (приоритет над WORK_UNITS)
        is_material = any(kw in name for kw in MATERIAL_KEYWORDS)
        if is_material:
            categories.append("Материал")
            continue
        
        # 3. Укрупнённые единицы — работа
        is_work_unit = any(wu in unit for wu in WORK_UNITS if wu not in LABOR_UNITS + ['%', 'компл', 'комплект'])
        if is_work_unit:
            categories.append("Работа")
            continue
        
        # 4. Проценты — накладные расходы/сметная прибыль
        if unit == '%' or 'накладные' in name or 'сметная прибыль' in name:
            categories.append("НР/СП")
            continue
        
        # 5. Проверка на заголовки/итоги по HEADER_PATTERNS
        is_header = any(pattern in name for pattern in HEADER_PATTERNS)
        if is_header:
            categories.append("Заголовок/Итог")
            continue
        
        # 6. Технические расчеты и объемы (TECHNICAL_PATTERNS) — Заголовок/Итог
        is_technical = any(name.startswith(pattern.rstrip(':')) for pattern in TECHNICAL_PATTERNS if pattern.endswith(':'))
        is_technical = is_technical or any(pattern in name for pattern in TECHNICAL_PATTERNS if not pattern.endswith(':'))
        # Дополнительно проверяем паттерны с двоеточием в конце как startswith
        for pattern in TECHNICAL_PATTERNS:
            if pattern.endswith(':'):
                clean_pattern = pattern.lstrip('^').rstrip(':')
                if name.startswith(clean_pattern):
                    is_technical = True
                    break
        if is_technical:
            categories.append("Заголовок/Итог")
            continue
        
        # 7. Проверка ключевых слов-процессов (работы) из WORK_KEYWORDS
        is_work = any(kw in name for kw in WORK_KEYWORDS)
        if is_work:
            categories.append("Работа")
            continue
        
        # 8. Проверка глагольных форм из VERB_PATTERNS — работа
        has_verb_pattern = any(verb in name for verb in VERB_PATTERNS)
        if has_verb_pattern:
            categories.append("Работа")
            continue
        
        # 9. MATERIAL_OVERRIDE — приоритет материала даже при наличии глаголов
        # (уже проверено выше, но оставляем для будущих расширений)
        
        # 10. Если единица измерения физическая (шт, м, м2, кг, т и т.д.)
        if any(pu in unit for pu in PHYSICAL_UNITS):
            # По умолчанию считаем материалом, если нет признаков работы
            categories.append("Материал (предпол.)")
            continue
        
        # 11. Все остальные случаи
        categories.append("Требует уточнения")
    
    result_df['Категория'] = categories
    return result_df


def filter_estimate_by_category(df: pd.DataFrame, category: str) -> pd.DataFrame:
    """
    Фильтрует позиции сметы по категории.
    
    Args:
        df: DataFrame с данными сметы (должна быть колонка "Категория")
        category: Категория для фильтрации ("Работа", "Материал", "НР/СП", etc.)
    
    Returns:
        Отфильтрованный DataFrame
    """
    if 'Категория' not in df.columns:
        df = classify_estimate_items(df)
    
    if category == "Материал":
        # Включаем как точные материалы, так и предположительные
        return df[df['Категория'].isin(['Материал', 'Материал (предпол.)'])]
    elif category == "Работа":
        return df[df['Категория'] == 'Работа']
    elif category == "НР/СП":
        return df[df['Категория'] == 'НР/СП']
    elif category == "Заголовок/Итог":
        return df[df['Категория'] == 'Заголовок/Итог']
    elif category == "Требует уточнения":
        return df[df['Категория'] == 'Требует уточнения']
    else:
        return df[df['Категория'] == category]


def get_estimate_summary(df: pd.DataFrame) -> dict:
    """
    Получает сводную статистику по классификации сметы.
    
    Args:
        df: DataFrame с данными сметы (должна быть колонка "Категория")
    
    Returns:
        Словарь со статистикой по категориям
    """
    if 'Категория' not in df.columns:
        df = classify_estimate_items(df)
    
    summary = df['Категория'].value_counts().to_dict()
    
    # Добавляем детализацию по материалам
    material_count = len(df[df['Категория'].isin(['Материал', 'Материал (предпол.)'])])
    work_count = len(df[df['Категория'] == 'Работа'])
    
    summary['Всего материалов'] = material_count
    summary['Всего работ'] = work_count
    summary['Всего позиций'] = len(df)
    
    return summary


def split_estimate_to_materials_and_works(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Разделяет смету на две таблицы: материалы и работы.
    
    Args:
        df: DataFrame с данными сметы
    
    Returns:
        Кортеж из двух DataFrame: (materials_df, works_df)
    """
    classified_df = classify_estimate_items(df)
    
    # Фильтруем материалы
    materials_df = filter_estimate_by_category(classified_df, "Материал")
    
    # Фильтруем работы
    works_df = filter_estimate_by_category(classified_df, "Работа")
    
    return materials_df, works_df


def calculate_totals_by_category(df: pd.DataFrame) -> dict:
    """
    Вычисляет суммы по категориям (материалы и работы).
    
    Args:
        df: DataFrame с данными сметы
    
    Returns:
        Словарь с суммами: {'materials_total': float, 'works_total': float}
    """
    classified_df = classify_estimate_items(df)
    
    materials_df = filter_estimate_by_category(classified_df, "Материал")
    works_df = filter_estimate_by_category(classified_df, "Работа")
    
    # Преобразуем колонку "Стоимость" в числовой формат
    def safe_to_numeric(series):
        return pd.to_numeric(series, errors='coerce').fillna(0)
    
    materials_total = safe_to_numeric(materials_df['Стоимость']).sum() if not materials_df.empty else 0
    works_total = safe_to_numeric(works_df['Стоимость']).sum() if not works_df.empty else 0
    
    return {
        'materials_total': materials_total,
        'works_total': works_total,
        'materials_count': len(materials_df),
        'works_count': len(works_df)
    }


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
    positions_df = extract_positions_by_sections(df)
    
    if positions_df.empty:
        return {
            'positions_total': 0.0,
            'positions_count': 0
        }
    
    # Преобразуем колонку "стоимость" в числовой формат
    positions_total = pd.to_numeric(positions_df['стоимость'], errors='coerce').fillna(0).sum()
    
    return {
        'positions_total': positions_total,
        'positions_count': len(positions_df)
    }



def delete_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Удаляет дубликаты и служебные строки из сметы после парсинга.
    
    Алгоритм очистки:
    1. Удаляет полные дубликаты строк (все колонки совпадают)
    2. Удаляет строки где все числовые колонки пусты или равны 0
    3. Удаляет технические/служебные строки типа "(Деревянные конструкции)"
    4. Удаляет строки с пустым полем "Ед.изм." (заголовки, итоги, служебные строки без единиц измерения)
    5. Удаляет строки с "ОТ(ЗТ)" где поле "Стоимость" пустое (дублируют цену основной строки)
    6. Удаляет дубликаты по ключевым полям: ['Обоснование', 'Наименование', 'Цена тек.', 'Стоимость']
       с сохранением первой записи
    
    Args:
        df: DataFrame с данными сметы после parse_estimate
    
    Returns:
        DataFrame очищенный от дубликатов и служебных строк
    """
    if df.empty:
        return df.copy()
    
    result_df = df.copy()
    
    # 1. Удаляем полные дубликаты строк
    result_df = result_df.drop_duplicates(keep='first')
    
    # 2. Фильтруем служебные строки и технические комментарии
    # Строки типа "(Деревянные конструкции)", "(...)" и т.п.
    if 'Наименование' in result_df.columns:
        mask_technical = result_df['Наименование'].astype(str).str.strip().str.startswith('(') & \
                        result_df['Наименование'].astype(str).str.strip().str.endswith(')')
        result_df = result_df[~mask_technical]
        
        # Также удаляем строки где Наименование пустое или содержит только пробелы
        result_df = result_df[result_df['Наименование'].astype(str).str.strip().isin(['', '-', 'NaN']) == False]
    
    # 3. Удаляем строки где все числовые колонки равны 0 или NaN
    numeric_cols = ['Кол-во на ед.', 'Коэф.', 'Кол-во всего', 'Цена баз.', 'Индекс', 'Цена тек.', 'Коэф.2', 'Стоимость']
    available_numeric_cols = [c for c in numeric_cols if c in result_df.columns]
    
    if available_numeric_cols:
        # Преобразуем числовые колонки к numeric
        for col in available_numeric_cols:
            result_df[col] = pd.to_numeric(result_df[col], errors='coerce')
        
        # Проверяем есть ли хотя бы одна ненулевая числовая колонка
        mask_nonzero = result_df[available_numeric_cols].apply(
            lambda row: row.notna().any() and (row != 0).any(), axis=1
        )
        result_df = result_df[mask_nonzero | result_df[available_numeric_cols].isna().all(axis=1)]
    
    # 4. Удаляем строки с пустым полем "Ед.изм." (заголовки, итоги, служебные строки)
    if 'Ед.изм.' in result_df.columns:
        # Считаем пустыми: NaN, пустую строку, '-', 'NaN'
        mask_empty_unit = result_df['Ед.изм.'].astype(str).str.strip().isin(['', 'nan', 'NaN', '-', 'None']) | \
                         result_df['Ед.изм.'].isna()
        result_df = result_df[~mask_empty_unit]
    
    # 5. Удаляем строки с "ОТ(ЗТ)" где поле "Стоимость" пустое или NaN
    # Такие строки дублируют цену основной строки и не несут полезной информации
    if 'Наименование' in result_df.columns and 'Стоимость' in result_df.columns:
        mask_ot_zt = result_df['Наименование'].astype(str).str.contains(r'ОТ\(ЗТ\)', regex=True, na=False)
        mask_empty_cost = result_df['Стоимость'].astype(str).str.strip().isin(['', 'NaN', 'nan', '-']) | \
                         result_df['Стоимость'].isna()
        # Преобразуем Стоимость в numeric для проверки на 0
        cost_numeric = pd.to_numeric(result_df['Стоимость'], errors='coerce')
        mask_zero_cost = cost_numeric.isna() | (cost_numeric == 0)
        
        # Удаляем строки где есть ОТ(ЗТ) И (Стоимость пустое ИЛИ Стоимость = 0)
        mask_delete_ot = mask_ot_zt | mask_zero_cost
        result_df = result_df[~mask_delete_ot]
    
    # 6. Удаляем дубликаты по ключевым полям: Обоснование + Наименование + Цена тек. + Стоимость
    key_columns = ['Обоснование', 'Наименование', 'Цена тек.', 'Стоимость']
    available_key_columns = [c for c in key_columns if c in result_df.columns]
    
    if len(available_key_columns) >= 2:  # Минимум 2 колонки для осмысленной проверки
        result_df = result_df.drop_duplicates(subset=available_key_columns, keep='first')
    
    # 7. Сбрасываем индексы
    result_df = result_df.reset_index(drop=True)
    
    return result_df


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
            
            # Применяем очистку от дубликатов к исходному DataFrame
            df_cleaned = delete_duplicates(df)
            
            # Извлекаем позиции по разделам
            positions_df = extract_positions_by_sections(df_cleaned)
            
            # Формируем имена листов с префиксом оригинального листа если их несколько
            prefix = f"{safe_sheet_name}_" if len(dfs) > 1 else ""
            
            # Лист: Позиции
            positions_sheet_name = f"{prefix}Позиции"
            if not positions_df.empty:
                positions_df.to_excel(writer, sheet_name=positions_sheet_name[:31], index=False)
            else:
                # Создаем пустой лист если нет позиций
                pd.DataFrame({'Наименование работ и затрат': ['Нет данных']}).to_excel(
                    writer, sheet_name=positions_sheet_name[:31], index=False
                )
    
    excel_content = excel_buffer.getvalue()
    excel_filename = f"{base_filename}_positions.xlsx"
    
    return excel_content, excel_filename
