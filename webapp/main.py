from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import io
import os

app = FastAPI(title="Smeta PWA")

# Подключаем статику (frontend)
app.mount("/static", StaticFiles(directory="webapp/static"), name="static")

@app.get("/")
async def root():
    """Отдаем главную страницу"""
    return FileResponse("webapp/index.html")

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
        df = pd.read_excel(io.BytesIO(contents), engine='openpyxl' if file.filename.endswith('.xlsx') else 'xlrd')

        # --- ЗДЕСЬ БУДЕТ ЛОГИКА ОБРАБОТКИ (как в телеграм боте) ---
        # Пока просто сохраняем как есть
        
        # Сохраняем результат в буфер CSV с кодировкой UTF-8
        stream = io.StringIO()
        # to_csv в StringIO уже работает с юникодом, но явно укажем, что мы работаем с текстом
        df.to_csv(stream, index=False, sep=';', encoding='utf-8-sig') 
        csv_content = stream.getvalue()

        # Возвращаем файл пользователю
        # utf-8-sig добавляет BOM, чтобы Excel корректно открывал кириллицу
        headers = {
            "Content-Disposition": f"attachment; filename*=UTF-8''processed_{file.filename.replace('.', '_')}.csv"
        }
        return StreamingResponse(iter([csv_content.encode('utf-8')]), media_type="text/csv", headers=headers)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки файла: {str(e)}")
