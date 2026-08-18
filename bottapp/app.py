import os
import io
import logging
import asyncio
import traceback
import tempfile
from typing import Any, Dict

from aiohttp import ClientTimeout
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart, Command
from aiogram.types import BufferedInputFile
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.client.telegram import TelegramAPIServer
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram_dialog import Dialog, StartMode, Window
from aiogram_dialog.widgets.kbd import Cancel, Start
from aiogram_dialog.widgets.text import Const, Format, Progress
from aiogram_dialog.widgets.input import TextInput, MessageInput

# Импорт бизнес-логики из core
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.estimates import parse_estimate as core_parse_estimate, export_estimates_to_csv

BOT_TOKEN = "8967391567:AAHa6VD74hzBiZhvvP3g62TiV0wZa5eLgxU"
PROXY_URL = "https://gentle-tree-1a2f.fln5kqj50.workers.dev"

# Если используете прокси, установите переменные окружения (опционально)
os.environ["HTTPS_PROXY"] = PROXY_URL
os.environ["HTTP_PROXY"] = PROXY_URL

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Увеличенные таймауты для сессии: total=300s, connect=30s, read=120s, write=120s
api_server = TelegramAPIServer.from_base(PROXY_URL)
session = AiohttpSession(
    api=api_server,
    timeout=ClientTimeout(total=300, connect=30, sock_read=120, sock_connect=30)
)
bot = Bot(token=BOT_TOKEN, session=session)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)


# --- FSM STATES ---
class EstimateState(StatesGroup):
    """Состояния для диалога обработки сметы"""
    waiting_for_file = State()
    processing = State()


# --- DIALOG HANDLERS ---
async def on_start(dialog, manager):
    """Обработчик начала диалога"""
    await manager.answer(
        "📄 **Загрузка сметы**\n\n"
        "Отправьте мне файл локальной сметы в формате `.xls` или `.xlsx`\n"
        "Максимальный размер: 20 МБ\n\n"
        "Или нажмите /cancel для отмены"
    )


async def on_file_received(message, widget, manager):
    """Обработчик получения файла"""
    document = message.document
    file_name = document.file_name
    
    # Проверка расширения
    valid_extensions = ('.xls', '.xlsx')
    if not any(file_name.lower().endswith(ext) for ext in valid_extensions):
        await message.answer(
            f"⚠️ Пожалуйста, загрузите файл в одном из поддерживаемых форматов: **{', '.join(valid_extensions)}**"
        )
        return
    
    # Проверка размера
    if document.file_size > 20 * 1024 * 1024:
        await message.answer("⚠️ Размер файла превышает 20 МБ. Пожалуйста, отправьте файл меньшего размера.")
        return
    
    # Сохраняем информацию о файле в state
    await manager.update(data={"file_id": document.file_id, "file_name": file_name})
    
    # Переключаемся на состояние обработки
    await manager.switch_to(EstimateState.processing)


async def process_estimate(dialog, manager):
    """Обработка сметы с прогресс-баром"""
    data = manager.get_data()
    file_id = data.get("file_id")
    file_name = data.get("file_name")
    
    if not file_id or not file_name:
        await manager.answer("❌ Ошибка: файл не найден")
        await manager.back()
        return
    
    progress_message = await manager.answer("⏳ Начинаю обработку...")
    logger.info(f"Начата обработка файла {file_name} от пользователя {manager.event.from_user.id}")
    
    try:
        # Шаг 1: Загрузка файла (25%)
        await progress_message.edit_text("⏳ Загрузка файла... [25%]")
        file = await bot.get_file(file_id, timeout=300)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_name)[1]) as tmp_file:
            tmp_path = tmp_file.name
            await bot.download_file(file.file_path, destination=tmp_file, timeout=300)
        
        logger.info(f"Файл {file_name} загружен во временное хранилище: {tmp_path}")
        
        # Шаг 2: Чтение файла (50%)
        await progress_message.edit_text("⏳ Чтение файла... [50%]")
        loop = asyncio.get_event_loop()
        file_content = await loop.run_in_executor(None, lambda: open(tmp_path, 'rb').read())
        os.unlink(tmp_path)
        file_stream = io.BytesIO(file_content)
        logger.info(f"Файл {file_name} загружен, размер: {len(file_content)} байт")
        
        # Шаг 3: Парсинг сметы (75%)
        await progress_message.edit_text("⏳ Парсинг сметы... [75%]")
        dfs = await loop.run_in_executor(None, core_parse_estimate, file_stream)
        
        if not dfs:
            await progress_message.edit_text("❌ Не удалось извлечь данные из сметы. Проверьте формат файла.")
            logger.warning(f"Пустой результат парсинга для файла {file_name}")
            await manager.back()
            return
        
        logger.info(f"Успешно распарсено {len(dfs)} листов из файла {file_name}")
        
        # Шаг 4: Экспорт в CSV (100%)
        await progress_message.edit_text("⏳ Экспорт в CSV... [100%]")
        base_filename = os.path.splitext(file_name)[0]
        result_content = await loop.run_in_executor(None, export_estimates_to_csv, dfs, base_filename)
        
        # Формируем имя и описание выходного файла
        if len(dfs) == 1:
            output_filename = f"{base_filename}_очищенная.csv"
            caption = (
                f"✅ Смета успешно очищена!\n\n"
                f"📊 Листов обработано: {len(dfs)}\n"
                f"📝 Строк данных: {len(dfs[0])}"
            )
        else:
            output_filename = f"{base_filename}_очищенная.zip"
            caption = (
                f"✅ Смета успешно очищена!\n\n"
                f"📊 Листов обработано: {len(dfs)}\n"
                f"📦 Файлы упакованы в ZIP архив"
            )
        
        # Отправляем результат
        await manager.answer_document(
            document=BufferedInputFile(result_content, filename=output_filename),
            caption=caption
        )
        
        await progress_message.delete()
        logger.info(f"Файл {output_filename} отправлен пользователю {manager.event.from_user.id}")
        
        # Завершаем диалог
        await manager.done()
        
    except asyncio.TimeoutError:
        logger.error(f"Timeout при загрузке файла {file_name}")
        await progress_message.edit_text(
            "❌ Превышено время загрузки файла. Попробуйте файл меньшего размера (до 20 МБ) или повторите позже."
        )
        await manager.back()
    except Exception as e:
        error_trace = traceback.format_exc()
        logger.error(f"Ошибка при обработке файла {file_name}: {e}\n{error_trace}")
        await progress_message.edit_text(
            f"❌ Произошла ошибка при обработке файла.\n\n"
            f"**Тип ошибки:** {type(e).__name__}\n"
            f"**Сообщение:** {str(e)}"
        )
        await manager.back()


async def on_close(dialog, manager):
    """Обработчик закрытия диалога"""
    await manager.clear_data()


# --- DIALOG CONFIGURATION ---
estimate_dialog = Dialog(
    Window(
        Const("📄 **Загрузка сметы**\n\nОтправьте мне файл локальной сметы в формате `.xls` или `.xlsx`\nМаксимальный размер: 20 МБ"),
        Start(
            Const("📁 Загрузить файл"),
            id="upload_file",
            state=EstimateState.waiting_for_file
        ),
        MessageInput(
            func=on_file_received,
            content_types=["document"]
        ),
        state=EstimateState.waiting_for_file,
    ),
    Window(
        Const("⏳ **Обработка файла**\n\nПожалуйста, дождитесь завершения обработки..."),
        Progress(
            field="progress",
            width=20,
            filled="█",
            empty="░"
        ),
        Format("⏳ Обработка: {percent:.0f}%"),
        state=EstimateState.processing,
    ),
    on_start=on_start,
    on_close=on_close,
)


# --- MESSAGE HANDLERS ---
@dp.message(CommandStart())
async def command_start_handler(message, dialog_manager):
    """Обработчик команды /start"""
    await dialog_manager.start(EstimateState.waiting_for_file, mode=StartMode.RESET_STACK)


@dp.message(Command("help"))
async def command_help_handler(message, dialog_manager):
    """Обработчик команды /help"""
    await message.answer(
        "ℹ️ **Справка по боту**\n\n"
        "Я умею:\n"
        "• Обрабатывать сметы в форматах .xls и .xlsx\n"
        "• Извлекать данные из таблиц смет\n"
        "• Экспортировать результаты в CSV формат\n\n"
        "📤 **Как использовать:**\n"
        "1. Нажмите /start для начала работы\n"
        "2. Отправьте файл сметы через интерфейс диалога\n"
        "3. Дождитесь обработки (отображается прогресс)\n"
        "4. Получите готовый CSV файл\n\n"
        "⚠️ **Важно:**\n"
        "• Максимальный размер файла: 20 МБ\n"
        "• Поддерживаются форматы .xls и .xlsx"
    )


@dp.message(Command("status"))
async def command_status_handler(message, dialog_manager):
    """Обработчик команды /status"""
    await message.answer("✅ Бот работает нормально и готов к обработке файлов.")


@dp.message(F.document)
async def handle_direct_document(message, dialog_manager):
    """Обработчик прямых загрузок файлов (вне диалога)"""
    # Если пользователь просто отправил файл без использования диалога
    document = message.document
    file_name = document.file_name
    
    valid_extensions = ('.xls', '.xlsx')
    if not any(file_name.lower().endswith(ext) for ext in valid_extensions):
        await message.answer(
            f"⚠️ Пожалуйста, загрузите файл в одном из поддерживаемых форматов: **{', '.join(valid_extensions)}**"
        )
        return
    
    if document.file_size > 20 * 1024 * 1024:
        await message.answer("⚠️ Размер файла превышает 20 МБ. Пожалуйста, отправьте файл меньшего размера.")
        return
    
    # Автоматически запускаем диалог обработки
    await dialog_manager.start(EstimateState.waiting_for_file, mode=StartMode.RESET_STACK)
    # Сохраняем файл в state и переключаемся на обработку
    await dialog_manager.update(data={"file_id": document.file_id, "file_name": file_name})
    await dialog_manager.switch_to(EstimateState.processing)


# --- РЕГИСТРАЦИЯ ДИАЛОГОВ ---
dp.include_router(estimate_dialog)


if __name__ == "__main__":
    try:
        # Запускаем polling без явного указания poll_timeout
        # Таймаут будет взят из настроек сессии (total=300s)
        asyncio.run(dp.start_polling(bot, allowed_updates=["message", "callback_query"]))
    except KeyboardInterrupt:
        print("Бот остановлен.")