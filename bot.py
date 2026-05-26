import os
from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes, ConversationHandler
from datetime import datetime
from openpyxl import Workbook
from database import add_car_record, get_today_salary, get_period_salary, get_all_cars_for_export, clear_today_records, clear_period_records, clear_all_records
import tempfile

# Загрузка переменных окружения
load_dotenv()

# Получение токена из переменных окружения
TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')

if not TOKEN:
    raise ValueError("TELEGRAM_BOT_TOKEN не найден в переменных окружения. Создайте файл .env и добавьте токен.")

# Состояния для ConversationHandler
WAITING_FOR_PHOTO, WAITING_FOR_CHECK = range(2)

# Хранилище временных данных
user_data = {}

# Функция для показа главного меню
async def show_main_menu(update):
    # Inline кнопки под сообщением (премиальный минимализм)
    keyboard = [
        [
            InlineKeyboardButton("▪ ЗП 1-15", callback_data='salary_1_15'),
            InlineKeyboardButton("▪ ЗП 16-30", callback_data='salary_16_30')
        ],
        [
            InlineKeyboardButton("▪ Сегодня", callback_data='today_salary'),
            InlineKeyboardButton("▪ Очистить ◻️◻️", callback_data='clear_menu')
        ],
        [
            InlineKeyboardButton("▪ Отчет 1-15", callback_data='export_1_15'),
            InlineKeyboardButton("▪ Отчет 16-30", callback_data='export_16_30')
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    if update.message:
        await update.message.reply_text(
            "<b>Car Wash</b>\n\n"
            "Выберите действие",
            parse_mode='HTML',
            reply_markup=reply_markup
        )
    else:
        await update.callback_query.message.reply_text(
            "<b>Car Wash</b>\n\n"
            "Выберите действие",
            parse_mode='HTML',
            reply_markup=reply_markup
        )

# Обработчик команды /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.effective_user
    
    await update.message.reply_text(
        f"<b>Car Wash</b>\n\n"
        f"Привет, {user.first_name}\n\n"
        "Инструкция:\n"
        "1. Отправьте фото машины\n"
        "2. Напишите сумму чека в описании\n"
        "3. Доля составит 25%\n\n"
        "Пример: фото с текстом \"5000\"",
        parse_mode='HTML'
    )
    
    await show_main_menu(update)

# Обработчик кнопок
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    
    if query.data == 'today_salary':
        salary = get_today_salary(user_id)
        await query.message.reply_text(
            "<b>Сегодня</b>\n\n"
            f"<code>{salary:.2f} ₽</code>",
            parse_mode='HTML'
        )
        await show_main_menu(update)
    
    elif query.data == 'salary_1_15':
        salary, cars = get_period_salary(user_id, 1, 15)
        await query.message.reply_text(
            "<b>Период 1-15</b>\n\n"
            f"Доля: <code>{salary:.2f} ₽</code>\n"
            f"Машин: <code>{len(cars)}</code>",
            parse_mode='HTML'
        )
        await show_main_menu(update)
    
    elif query.data == 'salary_16_30':
        salary, cars = get_period_salary(user_id, 16, 30)
        await query.message.reply_text(
            "<b>Период 16-30</b>\n\n"
            f"Доля: <code>{salary:.2f} ₽</code>\n"
            f"Машин: <code>{len(cars)}</code>",
            parse_mode='HTML'
        )
        await show_main_menu(update)
    
    elif query.data == 'export_1_15':
        await export_data(query, user_id, 1, 15)
    
    elif query.data == 'export_16_30':
        await export_data(query, user_id, 16, 30)
    
    elif query.data == 'clear_menu':
        keyboard = [
            [
                InlineKeyboardButton("▪ Очистить сегодня", callback_data='clear_today'),
                InlineKeyboardButton("▪ Очистить 1-15", callback_data='clear_1_15')
            ],
            [
                InlineKeyboardButton("▪ Очистить 16-30", callback_data='clear_16_30'),
                InlineKeyboardButton("▪ Очистить все", callback_data='clear_all')
            ],
            [InlineKeyboardButton("▪ Назад", callback_data='back')]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.message.reply_text(
            "<b>Очистка</b>\n\n"
            "Выберите период",
            parse_mode='HTML',
            reply_markup=reply_markup
        )
    
    elif query.data == 'clear_today':
        deleted = clear_today_records(user_id)
        await query.message.reply_text(
            "<b>Очистка</b>\n\n"
            f"Удалено: <code>{deleted}</code>",
            parse_mode='HTML'
        )
        await show_main_menu(query)
    
    elif query.data == 'clear_1_15':
        deleted = clear_period_records(user_id, 1, 15)
        await query.message.reply_text(
            "<b>Очистка</b>\n\n"
            f"Удалено: <code>{deleted}</code>",
            parse_mode='HTML'
        )
        await show_main_menu(query)
    
    elif query.data == 'clear_16_30':
        deleted = clear_period_records(user_id, 16, 30)
        await query.message.reply_text(
            "<b>Очистка</b>\n\n"
            f"Удалено: <code>{deleted}</code>",
            parse_mode='HTML'
        )
        await show_main_menu(query)
    
    elif query.data == 'clear_all':
        deleted = clear_all_records(user_id)
        await query.message.reply_text(
            "<b>Очистка</b>\n\n"
            f"Удалено: <code>{deleted}</code>",
            parse_mode='HTML'
        )
        await show_main_menu(query)
    
    elif query.data == 'back':
        await show_main_menu(query)

# Обработчик фото
async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    photo = update.message.photo[-1]
    caption = update.message.caption
    
    if caption:
        try:
            check_amount = float(caption.strip())
            if check_amount <= 0:
                await update.message.reply_text(
                    "<b>Ошибка</b>\n\n"
                    "Сумма должна быть положительной",
                    parse_mode='HTML'
                )
                return
            
            salary = add_car_record(user_id, photo_file_id=photo.file_id, check_amount=check_amount)
            
            await update.message.reply_text(
                "<b>Запись добавлена</b>\n\n"
                f"Чек: <code>{check_amount:.2f} ₽</code>\n"
                f"Доля: <code>{salary:.2f} ₽</code>",
                parse_mode='HTML'
            )
            
            await show_main_menu(update)
            
        except ValueError:
            await update.message.reply_text(
                "<b>Ошибка</b>\n\n"
                "Напишите число в описании",
                parse_mode='HTML'
            )
    else:
        user_data[user_id] = {'state': WAITING_FOR_CHECK, 'photo_file_id': photo.file_id}
        await update.message.reply_text(
            "<b>Фото получено</b>\n\n"
            "Отправьте сумму чека",
            parse_mode='HTML'
        )

# Обработчик текста
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    text = update.message.text
    
    if user_id in user_data and user_data[user_id].get('state') == WAITING_FOR_CHECK:
        try:
            check_amount = float(text)
            if check_amount <= 0:
                await update.message.reply_text(
                    "<b>Ошибка</b>\n\n"
                    "Сумма должна быть положительной",
                    parse_mode='HTML'
                )
                return
            
            photo_file_id = user_data[user_id].get('photo_file_id')
            salary = add_car_record(user_id, photo_file_id=photo_file_id, check_amount=check_amount)
            
            del user_data[user_id]
            
            await update.message.reply_text(
                "<b>Запись добавлена</b>\n\n"
                f"Чек: <code>{check_amount:.2f} ₽</code>\n"
                f"Доля: <code>{salary:.2f} ₽</code>",
                parse_mode='HTML'
            )
            
            await show_main_menu(update)
            
        except ValueError:
            await update.message.reply_text(
                "<b>Ошибка</b>\n\n"
                "Введите число",
                parse_mode='HTML'
            )

# Функция экспорта данных
async def export_data(query, user_id, start_day, end_day):
    cars_data = get_all_cars_for_export(user_id, start_day, end_day)
    
    if not cars_data:
        await query.message.reply_text(
            "<b>Нет данных</b>\n\n"
            "За этот период нет записей",
            parse_mode='HTML'
        )
        return
    
    # Считаем сумму чека и доли
    total_check = sum(car['check_amount'] for car in cars_data)
    total_salary = sum(car['salary_25_percent'] for car in cars_data)
    
    # Создаем Excel файл
    wb = Workbook()
    ws = wb.active
    ws.title = "Доля"
    
    ws.append(['Дата', 'Чек', 'Доля (25%)', 'Описание'])
    
    for car in cars_data:
        ws.append([
            car['date'],
            car['check_amount'],
            car['salary_25_percent'],
            car['description'] or ''
        ])
    
    with tempfile.NamedTemporaryFile(mode='wb', suffix='.xlsx', delete=False) as tmp:
        temp_filename = tmp.name
    
    wb.save(temp_filename)
    
    # Отправляем Excel файл с суммарными данными
    try:
        with open(temp_filename, 'rb') as file:
            await query.message.reply_document(
                document=file,
                caption=f"Таблица {start_day}-{end_day}\n\nСумма чека: {total_check:.2f} ₽\nДоля 25%: {total_salary:.2f} ₽"
            )
    except Exception as e:
        await query.message.reply_text(
            f"<b>Ошибка</b>\n\n"
            f"{e}",
            parse_mode='HTML'
        )
    finally:
        if os.path.exists(temp_filename):
            os.remove(temp_filename)
    
    # Отправляем фото
    await query.message.reply_text(
        f"<b>Фото за период {start_day}-{end_day}</b>\n\n"
        f"Сумма чека: {total_check:.2f} ₽\nДоля 25%: {total_salary:.2f} ₽",
        parse_mode='HTML'
    )
    
    for car in cars_data:
        if car['photo_file_id']:
            try:
                await query.message.reply_photo(
                    photo=car['photo_file_id'],
                    caption=f"{car['check_amount']} ₽ | {car['salary_25_percent']} ₽ | {car['date']}"
                )
            except Exception as e:
                await query.message.reply_text(
                    f"<b>Ошибка фото</b>\n\n"
                    f"{e}",
                    parse_mode='HTML'
                )
    
    await show_main_menu(query)

def main() -> None:
    application = Application.builder().token(TOKEN).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CallbackQueryHandler(button_handler))
    application.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    print("Бот запущен...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
