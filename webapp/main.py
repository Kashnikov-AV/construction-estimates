from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import io
import os

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
    Принимает Excel файл, обрабатывает и возвращает CSV.
    Поддерживает кириллицу в названиях колонок и содержимом.
    """
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Файл должен быть в формате .xlsx или .xls")

    try:
        # Читаем файл в память
        contents = await file.read()
        
        # Читаем Excel с явным указанием движка для поддержки xls/xlsx
        engine = 'openpyxl' if file.filename.endswith('.xlsx') else 'xlrd'
        
        try:
            df = pd.read_excel(io.BytesIO(contents), engine=engine)
        except Exception as e:
            # Если не получилось прочитать, пробуем без явного указания движка (иногда помогает)
            # или пробуем другой движок если первый не сработал
            if engine == 'xlrd':
                 # Пробуем openpyxl если xlrd не справился (редкий случай для старых xls)
                 df = pd.read_excel(io.BytesIO(contents), engine='openpyxl')
            else:
                 raise e

        # --- ЗДЕСЬ БУДЕТ ЛОГИКА ОБРАБОТКИ (как в телеграм боте) ---
        # Пока просто сохраняем как есть
        
        # Сохраняем результат в буфер CSV с кодировкой UTF-8
        stream = io.StringIO()
        # to_csv в StringIO уже работает с юникодом, но явно укажем, что мы работаем с текстом
        # utf-8-sig добавляет BOM, чтобы Excel корректно открывал кириллицу
        df.to_csv(stream, index=False, sep=';', encoding='utf-8-sig') 
        csv_content = stream.getvalue()

        # Возвращаем файл пользователю
        # Кодируем в UTF-8 перед отправкой
        csv_bytes = csv_content.encode('utf-8')
        
        headers = {
            "Content-Disposition": f"attachment; filename*=UTF-8''processed_{file.filename.replace('.', '_')}.csv",
            "Access-Control-Expose-Headers": "Content-Disposition"
        }
        return StreamingResponse(iter([csv_bytes]), media_type="text/csv; charset=utf-8", headers=headers)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки файла: {str(e)}")
