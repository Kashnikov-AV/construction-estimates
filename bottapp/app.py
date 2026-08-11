import os
import io
import logging
import asyncio
import aiohttp
import pandas as pd
import xlea

from aiohttp import ClientTimeout, TCPConnector
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.types import FSInputFile, BufferedInputFile
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.client.telegram import TelegramAPIServer
from aiogram.exceptions import TelegramNetworkError
from xlea import Schema, Column, config

BOT_TOKEN = "8967391567:AAHa6VD74hzBiZhvvP3g62TiV0wZa5eLgxU"
PROXY_URL = "https://gentle-tree-1a2f.fln5kqj50.workers.dev"   # если нужен прокси

# Если используете прокси, установите переменные окружения (опционально)
os.environ["HTTPS_PROXY"] = PROXY_URL
os.environ["HTTP_PROXY"] = PROXY_URL

dp = Dispatcher()
api_server = TelegramAPIServer.from_base(PROXY_URL)
session = AiohttpSession(api=api_server, timeout=120)   # без proxy
bot = Bot(token=BOT_TOKEN, session=session)

COLS = ["номер", "обоснование", "наименование", "ед_изм", "кол_на_ед", "коэф1",
        "всего_коэф", "база_ед", "индекс", "тек_ед", "коэф2", "всего_руб"]

def norm(v):
    t = str(v).strip()
    return t[:-2] if t.endswith(".0") else t

def money(v):
    if v is None or norm(v) == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace(" ", ""))
    except ValueError:
        return None

@config(header_rows=2)
class Row(Schema):
    c1:  str = Column(r"(^|;)1(?:\.0)?$",  regexp=True)
    c2:  str = Column(r"(^|;)2(?:\.0)?$",  regexp=True)
    c3:  str = Column(r"(^|;)3(?:\.0)?$",  regexp=True)
    c4:  str = Column(r"(^|;)4(?:\.0)?$",  regexp=True)
    c5:  str = Column(r"(^|;)5(?:\.0)?$",  regexp=True)
    c6:  str = Column(r"(^|;)6(?:\.0)?$",  regexp=True)
    c7:  str = Column(r"(^|;)7(?:\.0)?$",  regexp=True)
    c8:  str = Column(r"(^|;)8(?:\.0)?$",  regexp=True)
    c9:  str = Column(r"(^|;)9(?:\.0)?$",  regexp=True)
    c10: str = Column(r"(^|;)10(?:\.0)?$", regexp=True)
    c11: str = Column(r"(^|;)11(?:\.0)?$", regexp=True)
    c12: str = Column(r"(^|;)12(?:\.0)?$", regexp=True)

def parse_estimate(file_bytes: bytes) -> io.BytesIO:
    """
    Парсит смету из .xls (байты) с помощью xlea.
    Возвращает BytesIO с CSV-данными при успехе,
    либо BytesIO с текстом ошибки (начинается с "Ошибка:") при неудаче.
    """
    try:
        file_obj = io.BytesIO(file_bytes)
        rows = list(xlea.autoread(file_obj, schema=Row))
        df = pd.DataFrame([r.asdict() for r in rows])
        df["row_index"] = [r.row_index for r in rows]
        rename_map = {f"c{i+1}": name for i, name in enumerate(COLS)}
        df = df.rename(columns=rename_map)

        # Убираем пустые строки и строки с "Приложение №"
        df = df[~(df[COLS].apply(lambda r: all(norm(v) == "" for v in r), axis=1))].copy()
        df = df[~df.apply(lambda r: any("Приложение №" in norm(v) for v in r[COLS]), axis=1)].copy()

        # Преобразование чисел
        for c in ["кол_на_ед", "коэф1", "всего_коэф", "база_ед", "индекс", "тек_ед", "коэф2", "всего_руб"]:
            df[c] = df[c].map(money)
        for c in COLS[:4]:
            df[c] = df[c].map(norm)

        # Разделы и виды работ
        is_sec = df["номер"].str.startswith("Раздел", na=False)
        is_sub = df["номер"].str.contains("работы|водоотведение|Водоснабжение|санузла|Строительный мусор", na=False) & ~is_sec
        df["раздел"] = df["номер"].where(is_sec).ffill()
        df["вид_работ"] = df["номер"].where(is_sub).ffill()
        df["тип_работы"] = df["ед_изм"].map(
            lambda u: "услуги" if u in ("чел.-ч", "маш.-ч")
            else ("материалы" if u not in ("", "%", None) else "")
        )
        df["категория"] = "монтажные работы"

        csv_buffer = io.BytesIO()
        df.to_csv(csv_buffer, index=False, sep=';', encoding='utf-8-sig')
        csv_buffer.seek(0)
        return csv_buffer

    except Exception as e:
        # Формируем сообщение об ошибке
        error_msg = f"Error: \n{type(e).__name__}: {e}"
        # Возвращаем BytesIO с текстом ошибки
        error_buffer = io.BytesIO()
        error_buffer.write(error_msg.encode('utf-8'))
        error_buffer.seek(0)
        return error_buffer

# --- ОБРАБОТЧИКИ TELEGRAM ---

@dp.message(CommandStart())
async def command_start_handler(message: types.Message) -> None:
    await message.answer(
        "👋 Привет! Я бот для очистки строительных смет.\n\n"
        "📄 Просто отправьте мне файл локальной сметы в формате `.xls`, "
        "и я верну вам очищенный файл в формате `.csv`, готовый для анализа."
    )

@dp.message(F.document)
async def handle_document(message: types.Message) -> None:
    document = message.document
    file_name = document.file_name

    if not file_name.lower().endswith('.xls'):
        await message.answer("⚠️ Пожалуйста, загрузите файл в формате **.xls** (старый формат Excel).")
        return

    processing_msg = await message.answer("⏳ Загружаю и обрабатываю вашу смету...")

    try:
        # Скачиваем файл с увеличенным таймаутом (задано в сессии)
        file = await bot.get_file(document.file_id)
        file_bytes = await bot.download_file(file.file_path)
        file_content = file_bytes.read()

        # Парсим
        result_buffer = parse_estimate(file_content)
        result_content = result_buffer.getvalue()

        # Проверяем ошибку парсинга
        if result_content.startswith(b'Error:'):
            error_text = result_content.decode('utf-8')
            await processing_msg.edit_text(f"❌ {error_text}")
            return

        # Отправляем CSV
        output_filename = file_name.rsplit('.', 1)[0] + '_очищенная.csv'
        await message.answer_document(
            document=BufferedInputFile(result_content, filename=output_filename),
            caption="✅ Смета успешно очищена и переведена в CSV!"
        )
        await processing_msg.delete()

    except asyncio.TimeoutError:
        await processing_msg.edit_text(
            "❌ Превышено время загрузки файла. Попробуйте файл меньшего размера (до 20 МБ) или повторите позже."
        )
    except Exception as e:
        traceback.print_exc()  # полная ошибка в консоль
        await processing_msg.edit_text("❌ Произошла внутренняя ошибка при обработке файла.")


if __name__ == "__main__":
    try:
        asyncio.run(dp.start_polling(bot))
    except KeyboardInterrupt:
        print("Бот остановлен.")