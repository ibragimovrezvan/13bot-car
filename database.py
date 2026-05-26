import sqlite3
from datetime import datetime
import os

DATABASE_NAME = "car_wash.db"

def init_database():
    """Создание базы данных и таблицы для записей о машинах"""
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cars (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            photo_file_id TEXT,
            photo_path TEXT,
            check_amount REAL NOT NULL,
            salary_25_percent REAL NOT NULL,
            date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            description TEXT
        )
    ''')
    
    conn.commit()
    conn.close()

def add_car_record(user_id, photo_file_id=None, photo_path=None, check_amount=0, description=""):
    """Добавление записи о машине"""
    salary_25_percent = check_amount * 0.25
    
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        INSERT INTO cars (user_id, photo_file_id, photo_path, check_amount, salary_25_percent, description)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (user_id, photo_file_id, photo_path, check_amount, salary_25_percent, description))
    
    conn.commit()
    conn.close()
    
    return salary_25_percent

def get_today_salary(user_id):
    """Получение зарплаты за сегодня"""
    today = datetime.now().strftime('%Y-%m-%d')
    
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT SUM(salary_25_percent) FROM cars 
        WHERE user_id = ? AND DATE(date) = ?
    ''', (user_id, today))
    
    result = cursor.fetchone()
    conn.close()
    
    return result[0] if result[0] else 0

def get_period_salary(user_id, start_day, end_day, month=None, year=None):
    """Получение зарплаты за период (1-15 или 16-30)"""
    if month is None:
        month = datetime.now().month
    if year is None:
        year = datetime.now().year
    
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT SUM(salary_25_percent), photo_file_id, check_amount, date, description
        FROM cars 
        WHERE user_id = ? 
        AND strftime('%Y', date) = ?
        AND strftime('%m', date) = ?
        AND CAST(strftime('%d', date) AS INTEGER) BETWEEN ? AND ?
    ''', (user_id, str(year), str(month).zfill(2), start_day, end_day))
    
    result = cursor.fetchall()
    conn.close()
    
    total_salary = sum(row[0] for row in result if row[0]) if result else 0
    cars_data = []
    
    for row in result:
        if row[0]:
            cars_data.append({
                'salary': row[0],
                'photo_file_id': row[1],
                'check_amount': row[2],
                'date': row[3],
                'description': row[4]
            })
    
    return total_salary, cars_data

def get_all_cars_for_export(user_id, start_day, end_day, month=None, year=None):
    """Получение всех данных о машинах для экспорта"""
    if month is None:
        month = datetime.now().month
    if year is None:
        year = datetime.now().year
    
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT photo_file_id, check_amount, salary_25_percent, date, description
        FROM cars 
        WHERE user_id = ? 
        AND strftime('%Y', date) = ?
        AND strftime('%m', date) = ?
        AND CAST(strftime('%d', date) AS INTEGER) BETWEEN ? AND ?
        ORDER BY date ASC
    ''', (user_id, str(year), str(month).zfill(2), start_day, end_day))
    
    result = cursor.fetchall()
    conn.close()
    
    cars_data = []
    for row in result:
        cars_data.append({
            'photo_file_id': row[0],
            'check_amount': row[1],
            'salary_25_percent': row[2],
            'date': row[3],
            'description': row[4]
        })
    
    return cars_data

def get_month_stats(user_id, month=None, year=None):
    """Получение статистики за месяц"""
    if month is None:
        month = datetime.now().month
    if year is None:
        year = datetime.now().year
    
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT SUM(check_amount), SUM(salary_25_percent), COUNT(*)
        FROM cars 
        WHERE user_id = ? 
        AND strftime('%Y', date) = ?
        AND strftime('%m', date) = ?
    ''', (user_id, str(year), str(month).zfill(2)))
    
    result = cursor.fetchone()
    conn.close()
    
    total_check = result[0] if result[0] else 0
    total_salary = result[1] if result[1] else 0
    count = result[2] if result[2] else 0
    
    # Расчет среднего чека
    avg_check = total_check / count if count > 0 else 0
    
    return {
        'total_check': total_check,
        'total_salary': total_salary,
        'count': count,
        'avg_check': avg_check
    }

def clear_today_records(user_id):
    """Удаление всех записей за сегодня"""
    today = datetime.now().strftime('%Y-%m-%d')
    
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        DELETE FROM cars 
        WHERE user_id = ? AND DATE(date) = ?
    ''', (user_id, today))
    
    deleted_count = cursor.rowcount
    conn.commit()
    conn.close()
    
    return deleted_count

def clear_period_records(user_id, start_day, end_day, month=None, year=None):
    """Удаление записей за период"""
    if month is None:
        month = datetime.now().month
    if year is None:
        year = datetime.now().year
    
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        DELETE FROM cars 
        WHERE user_id = ? 
        AND strftime('%Y', date) = ?
        AND strftime('%m', date) = ?
        AND CAST(strftime('%d', date) AS INTEGER) BETWEEN ? AND ?
    ''', (user_id, str(year), str(month).zfill(2), start_day, end_day))
    
    deleted_count = cursor.rowcount
    conn.commit()
    conn.close()
    
    return deleted_count

def clear_all_records(user_id):
    """Удаление всех записей пользователя"""
    conn = sqlite3.connect(DATABASE_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        DELETE FROM cars 
        WHERE user_id = ?
    ''', (user_id,))
    
    deleted_count = cursor.rowcount
    conn.commit()
    conn.close()
    
    return deleted_count

# Инициализация базы данных при запуске
init_database()
