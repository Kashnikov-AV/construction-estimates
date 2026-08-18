from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import io
import os
import urllib.parse

# Импорт бизнес-логики из core
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.estimates import parse_estimate as core_parse_estimate, export_estimates_to_csv

app = FastAPI(title="Smeta PWA")

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
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Файл должен быть в формате .xlsx или .xls")

    try:
        # Читаем файл в память (BytesIO)
        contents = await file.read()
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
        
        # Кодируем имя файла для поддержки кириллицы (RFC 5987)
        encoded_filename = urllib.parse.quote(output_filename.encode('utf-8'))
        
        headers = {
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",
            "Access-Control-Expose-Headers": "Content-Disposition"
        }
        
        return StreamingResponse(iter([result]), media_type=media_type, headers=headers)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки файла: {str(e)}")
