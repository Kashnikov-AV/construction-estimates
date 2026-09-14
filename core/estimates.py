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


# Ключевые слова для классификации работ (глаголы и процессы)
WORK_KEYWORDS = [
    'монтаж', 'демонтаж', 'установка', 'устройство', 'укладка', 'окраска', 'сварка',
    'очистка', 'погрузка', 'разгрузка', 'разборка', 'снятие', 'прокладка', 'крепление',
    'испытание', 'наладка', 'регулировка', 'проверка', 'контроль', 'измерение',
    'подготовка', 'грунтовка', 'шпатлевка', 'штукатурка', 'бетонирование', 'армирование',
    'герметизация', 'изоляция', 'утепление', 'облицовка', 'футеровка', 'антикоррозийная',
    'огнезащита', 'маркировка', 'консервация', 'расконсервация', 'транспортирование',
    'перемещение', 'подъем', 'опускание', 'центрировка', 'выверка', 'юстировка',
    'опрессовка', 'продувка', 'промывка', 'дезинфекция', 'дегазация', 'обеззараживание'
]

# Ключевые слова для классификации материалов (существительные - ресурсы)
MATERIAL_KEYWORDS = [
    'бетон', 'раствор', 'смесь', 'цемент', 'песок', 'щебень', 'грунт', 'асфальт',
    'кирпич', 'блок', 'панель', 'плита', 'балка', 'колонна', 'ферма', 'каркас',
    'труба', 'трубопровод', 'кабель', 'провод', 'шина', 'арматура', 'сетка',
    'профиль', 'уголок', 'швеллер', 'двутавр', 'лист', 'полоса', 'круг', 'квадрат',
    'краска', 'лак', 'эмаль', 'грунтовка', 'шпатлевка', 'штукатурка', 'клей',
    'герметик', 'мастика', 'изоляцион', 'утеплител', 'рубероид', 'толь', 'пергамин',
    'стекло', 'зеркало', 'плинтус', 'наличник', 'дверь', 'ворота', 'калитка',
    'окно', 'витраж', 'люк', 'решетка', 'ограждение', 'забор', 'ворота',
    'лампа', 'светильник', 'прожектор', 'трансформатор', 'генератор', 'двигатель',
    'насос', 'вентилятор', 'компрессор', 'котел', 'бойлер', 'радиатор', 'конвектор',
    'шкаф', 'щит', 'панель', 'пульт', 'прибор', 'датчик', 'сенсор', 'счетчик',
    'задвижка', 'клапан', 'кран', 'вентиль', 'фильтр', 'редуктор', 'муфта',
    'болт', 'гайка', 'винт', 'шуруп', 'саморез', 'заклепка', 'дюбель', 'анкер',
    'электрод', 'проволока', 'лента', 'скотч', 'изолента', 'термоусадка',
    'бензин', 'дизель', 'керосин', 'масло', 'смазка', 'топливо', 'электроэнергия'
]

# Единицы измерения, характерные для работ (укрупненные)
WORK_UNITS = [
    '100 м2', '100 м3', '100 м', '100 шт', '10 м2', '10 м3', '10 м', '10 шт',
    '1000 шт', '10 компл', '100 отверстий', '100 соединений', '100 м ступеней',
    'чел.-ч', 'маш.-ч', '%', 'компл'
]


def classify_estimate_items(df: pd.DataFrame) -> pd.DataFrame:
    """
    Классифицирует позиции сметы на работы/услуги и материалы.
    
    Алгоритм классификации:
    1. Единицы измерения труда (чел.-ч, маш.-ч) — работа
    2. Укрупнённые единицы (100 м2, 10 шт и т.п.) — работа
    3. Проценты (%) — накладные расходы/сметная прибыль (не материал)
    4. Ключевые слова-процессы в наименовании — работа
    5. Ключевые слова-материалы в наименовании — материал
    6. Физические единицы без маркеров — требует проверки по наименованию
    
    Args:
        df: DataFrame с данными сметы (колонки: "Наименование", "Ед.изм.")
    
    Returns:
        DataFrame с добавленной колонкой "Категория" со значениями:
        - "Работа" — работы и услуги
        - "Материал" — материальные ресурсы
        - "НР/СП" — накладные расходы и сметная прибыль
        - "Заголовок/Итог" — служебные строки
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
        if unit in ['чел.-ч', 'маш.-ч']:
            categories.append("Работа")
            continue
        
        # 2. Укрупнённые единицы — работа
        is_work_unit = any(wu in unit for wu in WORK_UNITS if wu not in ['чел.-ч', 'маш.-ч', '%', 'компл'])
        if is_work_unit:
            categories.append("Работа")
            continue
        
        # 3. Проценты — накладные расходы/сметная прибыль
        if unit == '%' or 'накладные' in name or 'сметная прибыль' in name:
            categories.append("НР/СП")
            continue
        
        # 4. Проверка ключевых слов-процессов (работы)
        is_work = any(kw in name for kw in WORK_KEYWORDS)
        if is_work:
            categories.append("Работа")
            continue
        
        # 5. Проверка ключевых слов-материалов
        is_material = any(kw in name for kw in MATERIAL_KEYWORDS)
        if is_material:
            categories.append("Материал")
            continue
        
        # 6. Физические единицы без явных маркеров — требует уточнения
        # Но если есть глагол в любой форме — это работа
        has_verb_pattern = any(verb in name for verb in [
            'устройств', 'монтаж', 'демонтаж', 'установк', 'укладк', 'окраск',
            'сварк', 'очистк', 'погрузк', 'разгрузк', 'разборк', 'снят', 'прокладк',
            'креплен', 'испытан', 'налаж', 'регулировк', 'проверк', 'измерен',
            'подготовк', 'грунтовк', 'шпатлевк', 'штукатурк', 'бетонирован', 'армирован',
            'герметиз', 'изолирован', 'утеплен', 'облицовк', 'футеровк', 'антикорроз',
            'огнезащит', 'маркировк', 'консервац', 'расконсервац', 'транспортирован',
            'перемещен', 'подъем', 'опускан', 'центрировк', 'выверк', 'юстировк',
            'опрессовк', 'продувк', 'промывк', 'дезинфекц', 'дегазац', 'обеззараживан'
        ])
        
        if has_verb_pattern:
            categories.append("Работа")
            continue
        
        # Если единица измерения физическая (шт, м, м2, кг, т и т.д.)
        physical_units = ['шт', 'м', 'м2', 'м3', 'кг', 'т', 'л', 'кВт·ч', 'кВт-ч']
        if any(pu in unit for pu in physical_units):
            # По умолчанию считаем материалом, если нет признаков работы
            categories.append("Материал (предпол.)")
            continue
        
        # Все остальные случаи
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
