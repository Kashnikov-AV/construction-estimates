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
    Принимает Excel файл, обрабатывает (пока пусто) и возвращает CSV.
    """
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Файл должен быть в формате .xlsx или .xls")

    try:
        # Читаем файл в память
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))

        # --- ЗДЕСЬ БУДЕТ ЛОГИКА ОБРАБОТКИ (как в телеграм боте) ---
        # Пока просто сохраняем как есть или создаем пустой файл по требованию
        
        # Если нужно вернуть именно "пустой" файл как в запросе (но с правильной структурой):
        # df = pd.DataFrame() 
        
        # Сохраняем результат в буфер CSV
        stream = io.StringIO()
        df.to_csv(stream, index=False, sep=';') # Используем точку с запятой для Excel совместимости
        csv_content = stream.getvalue()

        # Возвращаем файл пользователю
        headers = {
            "Content-Disposition": f"attachment; filename=processed_{file.filename.replace('.', '_')}.csv"
        }
        return StreamingResponse(iter([csv_content]), media_type="text/csv", headers=headers)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки файла: {str(e)}")
