from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import io
import os
import urllib.parse
import re

# Импорт бизнес-логики из core
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.estimates import parse_estimate as core_parse_estimate, export_estimates_to_csv

app = FastAPI(title="Smeta PWA")

# Таблица транслитерации для кириллицы
CYRILLIC_TO_LATIN = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo', 'ж': 'zh',
    'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o',
    'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'kh', 'ц': 'ts',
    'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu',
    'я': 'ya',
    'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'Ё': 'Yo', 'Ж': 'Zh',
    'З': 'Z', 'И': 'I', 'Й': 'Y', 'К': 'K', 'Л': 'L', 'М': 'M', 'Н': 'N', 'О': 'O',
    'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U', 'Ф': 'F', 'Х': 'Kh', 'Ц': 'Ts',
    'Ч': 'Ch', 'Ш': 'Sh', 'Щ': 'Sch', 'Ъ': '', 'Ы': 'Y', 'Ь': '', 'Э': 'E', 'Ю': 'Yu',
    'Я': 'Ya'
}

def transliterate(text: str) -> str:
    """Транслитерирует кириллический текст в латиницу и удаляет недопустимые символы для имен файлов."""
    result = []
    for char in text:
        # Транслитерируем кириллицу
        if char in CYRILLIC_TO_LATIN:
            result.append(CYRILLIC_TO_LATIN[char])
        # Оставляем только безопасные ASCII символы (буквы, цифры, точка, дефис, подчеркивание)
        elif ord(char) < 128 and (char.isalnum() or char in '._-'):
            result.append(char)
        # Все остальные символы (включая №, пробелы, спецсимволы) заменяем на подчеркивание или пропускаем
        elif char == ' ':
            result.append('_')
        # Остальные недопустимые символы просто пропускаем
    return ''.join(result)

# Определяем базовую директорию проекта
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
INDEX_HTML = os.path.join(BASE_DIR, "index.html")

# Подключаем статику (frontend)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/")
async def root():
    """Отдаем главную страницу"""
    return FileResponse(INDEX_HTML, media_type="text/html")

@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
    """
    Принимает Excel файл, обрабатывает и возвращает CSV (или ZIP с CSV).
    Поддерживает кириллицу в названиях колонок и содержимом.
    Если файл содержит несколько листов, возвращается ZIP-архив.
    """
    # Проверяем расширение файла (для совместимости)
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Файл должен быть в формате .xlsx или .xls")

    try:
        # Читаем файл в память (BytesIO)
        contents = await file.read()
        
        # Проверяем сигнатуру файла для определения реального формата
        # .xlsx (Office Open XML) начинается с PK (504b0304)
        # .xls (BIFF8) начинается с d0cf11e0a1b11ae1
        if len(contents) < 8:
            raise HTTPException(status_code=400, detail="Файл слишком мал или поврежден")
        
        header = contents[:8]
        is_xls = header.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')
        is_xlsx = header.startswith(b'PK\x03\x04')
        
        if not is_xls and not is_xlsx:
            raise HTTPException(
                status_code=400, 
                detail=f"Неподдерживаемый формат файла. Ожидался .xls или .xlsx, но файл не является корректным Excel файлом."
            )
        
        file_stream = io.BytesIO(contents)
        
        # Обрабатываем файл через функцию парсинга напрямую из потока
        dfs = core_parse_estimate(file_stream)
        
        if not dfs:
            raise HTTPException(status_code=400, detail="Не удалось извлечь данные из файла")
        
        # Получаем базовое имя файла без расширения
        base_filename = os.path.splitext(file.filename)[0]
        
        # Экспортируем в CSV (или ZIP если несколько листов)
        result = export_estimates_to_csv(dfs, base_filename)
        
        # Определяем тип контента и имя файла
        if len(dfs) == 1:
            # Один лист - возвращаем CSV
            media_type = "text/csv; charset=utf-8"
            output_filename = f"{base_filename}.csv"
        else:
            # Несколько листов - возвращаем ZIP
            media_type = "application/zip"
            output_filename = f"{base_filename}.zip"
        
        # Кодируем имя файла для поддержки кириллицы через транслитерацию
        # Используем только транслитерированное имя для максимальной совместимости со всеми браузерами
        translit_filename = transliterate(output_filename)
        
        headers = {
            "Content-Disposition": f"attachment; filename=\"{translit_filename}\"",
            "Access-Control-Expose-Headers": "Content-Disposition"
        }
        
        return StreamingResponse(iter([result]), media_type=media_type, headers=headers)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки файла: {str(e)}")
