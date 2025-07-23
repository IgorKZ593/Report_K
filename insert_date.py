# insert_date.py

# Проверка и импорт необходимых модулей
def check_modules():
    """
    Проверяет наличие необходимых модулей и уведомляет пользователя, если какой-либо отсутствует.
    """
    import importlib
    import sys
    required = ['holidays', 'datetime', 'json', 'keyboard', 'os', 'sys']
    for mod in required:
        try:
            importlib.import_module(mod)
        except ImportError:
            print(f"Модуль '{mod}' не установлен. Установите его командой: pip install {mod}")
            sys.exit(1)

check_modules()

import sys
import os
import json
import datetime
import holidays
import keyboard

def print_welcome():
    """
    Выводит приветственное сообщение.
    """
    print("Подготовка аналитического отчета для клиентов N1 Broker")

def wait_for_esc():
    """
    Проверяет, нажата ли клавиша Esc для выхода.
    """
    print("Для выхода нажмите Esc")
    if keyboard.is_pressed('esc'):
        print("Выход по Esc.")
        sys.exit(0)

def is_weekend(date_obj):
    """
    Проверяет, является ли дата выходным (суббота или воскресенье).
    """
    return date_obj.weekday() >= 5

def is_us_holiday(date_obj, holidays_us):
    """
    Проверяет, является ли дата национальным праздником США.
    """
    return date_obj in holidays_us

def find_nearest_valid_dates(date_obj, min_date, holidays_us):
    """
    Находит ближайшие доступные даты до и после указанной.
    """
    before = date_obj - datetime.timedelta(days=1)
    after = date_obj + datetime.timedelta(days=1)
    # Ищем предыдущую доступную дату
    while before >= min_date and (is_weekend(before) or is_us_holiday(before, holidays_us)):
        before -= datetime.timedelta(days=1)
    # Ищем следующую доступную дату
    while is_weekend(after) or is_us_holiday(after, holidays_us):
        after += datetime.timedelta(days=1)
    return before, after

def get_date_input(prompt, min_date, holidays_us, start_date=None):
    """
    Запрашивает у пользователя дату, выполняет все проверки и возвращает объект datetime.
    """
    while True:
        date_str = input(prompt)
        wait_for_esc()
        try:
            date_obj = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
        except ValueError:
            print("Неверный формат даты. Используйте формат dd/mm/yyyy.")
            continue

        if date_obj < min_date:
            print("Указанный вами период не может быть применен, так как выходит за период деятельности N1 Broker")
            continue

        if start_date and date_obj <= start_date:
            print("Неверный диапазон дат. Конечная дата должна быть старше начальной")
            continue

        if is_weekend(date_obj):
            print(f"{date_obj.strftime('%d.%m.%Y')} — выходной день")
            before, after = find_nearest_valid_dates(date_obj, min_date, holidays_us)
            print(f"Ближайшие доступные даты для формирования отчета: {before.strftime('%d.%m.%Y')}, {after.strftime('%d.%m.%Y')}")
            continue

        if is_us_holiday(date_obj, holidays_us):
            holiday_name = holidays_us.get(date_obj)
            print(f"{date_obj.strftime('%d.%m.%Y')} — {holiday_name}")
            before, after = find_nearest_valid_dates(date_obj, min_date, holidays_us)
            print(f"Ближайшие доступные даты для формирования отчета: {before.strftime('%d.%m.%Y')}, {after.strftime('%d.%m.%Y')}")
            continue

        return date_obj

def save_dates_to_json(start_date, end_date, path):
    """
    Сохраняет даты в JSON-файл по указанному пути.
    """
    data = {
        "start_date": start_date.strftime("%d.%m.%Y"),
        "end_date": end_date.strftime("%d.%m.%Y")
    }
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"Даты сохранены в {path}")

def main():
    """
    Основная функция скрипта.
    """
    print_welcome()
    min_date = datetime.date(2022, 1, 1)
    holidays_us = holidays.US(years=range(2022, datetime.date.today().year + 2))

    # Ввод начальной даты
    start_date = get_date_input("Введите дату начала отчета (dd/mm/yyyy): ", min_date, holidays_us)
    print(f"Дата начала: {start_date.strftime('%d.%m.%Y')}")
    wait_for_esc()

    # Ввод конечной даты
    end_date = get_date_input("Введите дату завершения отчета (dd/mm/yyyy): ", min_date, holidays_us, start_date=start_date)
    print(f"Дата завершения отчета: {end_date.strftime('%d.%m.%Y')}")
    wait_for_esc()

    # Сохраняем в JSON
    save_path = os.path.join("Data_work", "report_dates.json")
    save_dates_to_json(start_date, end_date, save_path)

if __name__ == "__main__":
    main()
