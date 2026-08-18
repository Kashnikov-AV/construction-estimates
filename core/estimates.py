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


KEYS = ["№ п/п", "Обоснование", "Наименование работ и затрат"]
COLS = ["№", "Обоснование", "Наименование", "Ед.изм.", "Кол-во на ед.", "Коэф.",
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
        
        # Проверяем сигнатуру файла: xls начинается с D0 CF 11 E0 A1 B1 1A E1
        if header.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
            xf = pd.ExcelFile(file_input, engine='xlrd')
        else:
            xf = pd.ExcelFile(file_input, engine='openpyxl')
    
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
        df = pd.DataFrame(rows, columns=COLS)
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
        df.to_csv(csv_buffer, index=False, sep=';', encoding='utf-8-sig')
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
            df.to_csv(csv_buffer, index=False, sep=';', encoding='utf-8-sig')
            csv_content = csv_buffer.getvalue().encode('utf-8-sig')
            
            zip_file.writestr(csv_filename, csv_content)
    
    return zip_buffer.getvalue()
