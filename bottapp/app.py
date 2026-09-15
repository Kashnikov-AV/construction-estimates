import os
import io
import logging
import asyncio
import traceback
import zipfile
from dotenv import load_dotenv

# Загружаем переменные окружения из файла .env
load_dotenv()

from aiohttp import ClientTimeout, TCPConnector
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart, Command
from aiogram.types import FSInputFile, BufferedInputFile
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.client.telegram import TelegramAPIServer
from aiogram.exceptions import TelegramNetworkError

# Импорт бизнес-логики из core
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.estimates import (
    parse_estimate as core_parse_estimate,
    export_estimates_to_csv,
    calculate_totals_by_category,
    export_estimate_to_excel,
    delete_duplicates
)

# Получаем секретные данные из переменных окружения
BOT_TOKEN = os.getenv("BOT_TOKEN")
PROXY_URL = os.getenv("PROXY_URL", "https://gentle-tree-1a2f.fln5kqj50.workers.dev")

if not BOT_TOKEN:
    raise ValueError("BOT_TOKEN не найден в переменных окружения. Пожалуйста, создайте файл .env с переменной BOT_TOKEN.")

# Если используете прокси, установите переменные окружения (опционально)
os.environ["HTTPS_PROXY"] = PROXY_URL
os.environ["HTTP_PROXY"] = PROXY_URL

dp = Dispatcher()
api_server = TelegramAPIServer.from_base(PROXY_URL)
# Увеличенные таймауты для сессии: connect=30s, read=120s, write=120s, pool=60s
session = AiohttpSession(
    api=api_server,
    timeout=60
)
bot = Bot(token=BOT_TOKEN, session=session)

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- ОБРАБОТЧИКИ TELEGRAM ---

@dp.message(CommandStart())
async def command_start_handler(message: types.Message) -> None:
    """Обработчик команды /start"""
    await message.answer(
        "👋 Привет! Я бот для очистки строительных смет.\n\n"
        "📋 Что я умею:\n"
        "• Автоматически определять формат сметы (.xls или .xlsx)\n"
        "• Извлекать данные из таблиц с любым расположением колонок\n"
        "• Очищать данные от лишних строк и форматирования\n"
        "• Экспортировать результат в удобный CSV формат\n\n"
        "📄 Как использовать:\n"
        "1. Отправьте мне файл локальной сметы в формате `.xls` или `.xlsx`\n"
        "2. Я обработаю файл и извлеку все данные\n"
        "3. Вы получите очищенный CSV файл (или ZIP при нескольких листах)\n\n"
        "🔧 Команды:\n"
        "/help - подробная справка\n"
        "/status - проверка работоспособности"
    )


@dp.message(Command("help"))
async def command_help_handler(message: types.Message) -> None:
    """Обработчик команды /help"""
    await message.answer(
        "ℹ️ Справка по боту\n\n"
        "Я умею:\n"
        "• Обрабатывать сметы в форматах .xls и .xlsx\n"
        "• Извлекать данные из таблиц смет\n"
        "• Экспортировать результаты в CSV формат\n\n"
        "📤 Как использовать:\n"
        "1. Отправьте мне файл сметы\n"
        "2. Дождитесь обработки\n"
        "3. Получите готовый CSV файл\n\n"
        "⚠️ Важно:\n"
        "• Максимальный размер файла: 20 МБ\n"
        "• Поддерживаются форматы .xls и .xlsx"
    )


@dp.message(Command("status"))
async def command_status_handler(message: types.Message) -> None:
    """Обработчик команды /status"""
    await message.answer("✅ Бот работает нормально и готов к обработке файлов.")


@dp.message(F.document)
async def handle_document(message: types.Message) -> None:
    """Обработчик загруженных файлов"""
    document = message.document
    file_name = document.file_name

    # Проверка расширения файла
    valid_extensions = ('.xls', '.xlsx')
    if not any(file_name.lower().endswith(ext) for ext in valid_extensions):
        await message.answer(
            f"⚠️ Пожалуйста, загрузите файл в одном из поддерживаемых форматов: {', '.join(valid_extensions)}"
        )
        return

    # Проверка размера файла (20 МБ лимит Telegram)
    if document.file_size > 20 * 1024 * 1024:
        await message.answer("⚠️ Размер файла превышает 20 МБ. Пожалуйста, отправьте файл меньшего размера.")
        return

    processing_msg = await message.answer(
        "⏳ Файл получен!\n\n"
        f"📁 Имя файла: `{file_name}`\n"
        "🔄 Начинаю загрузку и обработку...\n\n"
        "Это может занять несколько секунд в зависимости от размера файла."
    )
    logger.info(f"Получен файл от пользователя {message.from_user.id}: {file_name}")

    try:
        # Асинхронная загрузка файла с явным указанием таймаута
        file = await bot.get_file(document.file_id, request_timeout=300)
        
        # Создаем временный файл для асинхронной загрузки больших файлов
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_name)[1]) as tmp_file:
            tmp_path = tmp_file.name
            # Асинхронное скачивание файла напрямую во временный файл
            await bot.download_file(file.file_path, destination=tmp_file, timeout=300)
        
        logger.info(f"Файл {file_name} загружен во временное хранилище: {tmp_path}")

        # Читаем файл асинхронно
        loop = asyncio.get_event_loop()
        file_content = await loop.run_in_executor(None, lambda: open(tmp_path, 'rb').read())
        
        # Удаляем временный файл
        os.unlink(tmp_path)
        
        file_stream = io.BytesIO(file_content)
        logger.info(f"Файл {file_name} загружен, размер: {len(file_content)} байт")

        # Парсим смету с помощью функции из core (запускаем в executor для неблокирующего выполнения)
        dfs = await loop.run_in_executor(None, core_parse_estimate, file_stream)

        if not dfs:
            await processing_msg.edit_text("❌ Не удалось извлечь данные из сметы. Проверьте формат файла.")
            logger.warning(f"Пустой результат парсинга для файла {file_name}")
            return

        logger.info(f"Успешно распарсено {len(dfs)} листов из файла {file_name}")

        # Применяем очистку от дубликатов к каждому листу
        dfs_cleaned = []
        for i, df in enumerate(dfs):
            df_cleaned = await loop.run_in_executor(None, delete_duplicates, df)
            dfs_cleaned.append(df_cleaned)
        
        # Используем очищенные DataFrame для дальнейшей обработки
        dfs = dfs_cleaned
        
        # Вычисляем суммы по материалам и работам для каждого листа
        totals_info = []
        for i, df in enumerate(dfs):
            sheet_name = df.attrs.get("sheet", f"Лист {i+1}")
            totals = await loop.run_in_executor(None, calculate_totals_by_category, df)
            totals_info.append({
                'sheet': sheet_name,
                'materials_total': totals['materials_total'],
                'works_total': totals['works_total'],
                'materials_count': totals['materials_count'],
                'works_count': totals['works_count']
            })

        # Формируем сообщение с суммами
        summary_message = "📊 Сводка по смете:\n\n"
        for info in totals_info:
            summary_message += f"📄 {info['sheet']}:\n"
            summary_message += f"   🔹 Материалы: {info['materials_count']} поз. на сумму {info['materials_total']:,.2f} руб.\n"
            summary_message += f"   🔹 Работы: {info['works_count']} поз. на сумму {info['works_total']:,.2f} руб.\n"
            summary_message += f"   💰 Итого: {info['materials_total'] + info['works_total']:,.2f} руб.\n\n"

        # Отправляем сообщение с суммами
        await message.answer(summary_message)

        # Экспортируем в Excel с разделением на материалы и работы
        base_filename = os.path.splitext(file_name)[0]
        excel_content, output_filename = await loop.run_in_executor(None, export_estimate_to_excel, dfs, base_filename)

        caption = "✅ Excel файл с разделением на Материалы и Работы готов!"

        # Отправляем Excel файл
        await message.answer_document(
            document=BufferedInputFile(excel_content, filename=output_filename),
            caption=caption
        )
        
        await processing_msg.delete()
        logger.info(f"Файл {output_filename} отправлен пользователю {message.from_user.id}")


    except asyncio.TimeoutError:
        logger.error(f"Timeout при загрузке файла {file_name} от пользователя {message.from_user.id}")
        await processing_msg.edit_text(
            "❌ Превышено время загрузки файла. Попробуйте файл меньшего размера (до 20 МБ) или повторите позже."
        )
    except Exception as e:
        error_trace = traceback.format_exc()
        logger.error(f"Ошибка при обработке файла {file_name}: {e}\n{error_trace}")
        await processing_msg.edit_text(
            f"❌ Произошла ошибка при обработке файла.\n\n"
            f"Тип ошибки: {type(e).__name__}\n"
            f"Сообщение: {str(e)}\n\n"
            f"Попробуйте отправить файл другого формата или обратитесь к разработчику."
        )


if __name__ == "__main__":
    try:
        asyncio.run(dp.start_polling(bot))
    except KeyboardInterrupt:
        print("Бот остановлен.")