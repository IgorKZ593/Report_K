# insert_date.py
"""
Требуется библиотека rich для цветного вывода.
Установить вручную при необходимости: pip install rich
"""

import os
import sys
import json

# Автоматическая установка rich, если не установлен
try:
    from rich import print
except ImportError:
    os.system(f"{sys.executable} -m pip install rich")
    from rich import print

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
            print(f"[bold red]Модуль '{mod}' не установлен. Установите его командой: pip install {mod}[/bold red]")
            sys.exit(1)

check_modules()

import datetime
import holidays
import keyboard

# Проверка и удаление существующего файла report_dates.json
json_path = os.path.join("Data_work", "report_dates.json")
if os.path.exists(json_path):
    try:
        os.remove(json_path)
    except Exception:
        pass  # Не выводим ошибку, если не удалось удалить

def print_welcome():
    """
    Выводит приветственное сообщение.
    """
    print("[bold green]Подготовка аналитического отчета для клиентов N1 Broker[/bold green]")

def wait_for_esc():
    """
    Проверяет, нажата ли клавиша Esc для выхода.
    """
    print("[bold yellow]Для выхода нажмите Esc[/bold yellow]")
    if keyboard.is_pressed('esc'):
        print("[bold red]Выход по Esc.[/bold red]")
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

def suggest_previous_valid_date(date_obj, min_date, holidays_us):
    """
    Возвращает ближайшую допустимую дату до date_obj (не выходной и не праздник), не раньше min_date.
    """
    prev_date = date_obj - datetime.timedelta(days=1)
    while prev_date >= min_date:
        if not is_weekend(prev_date) and not is_us_holiday(prev_date, holidays_us):
            return prev_date
        prev_date -= datetime.timedelta(days=1)
    return None

def get_date_input(prompt, min_date, holidays_us, start_date=None):
    """
    Запрашивает у пользователя дату, выполняет все проверки и возвращает объект datetime.
    Добавлена логика проверки на сегодня и будущее, с подбором ближайшей допустимой даты назад.
    """
    while True:
        date_str = input(prompt)
        wait_for_esc()
        try:
            date_obj = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
        except ValueError:
            print("[bold red]Ошибка: введена некорректная дата! Используйте dd/mm/yyyy.[/bold red]")
            continue

        today = datetime.date.today()

        # --- Новый блок: Проверка на сегодня и будущее ---
        if date_obj == today:
            print("[bold red]Невозможно сделать отчет на текущую дату[/bold red]")
            suggested = suggest_previous_valid_date(date_obj, min_date, holidays_us)
            if suggested:
                print(f"[bold yellow]Предлагаемая дата: {suggested.strftime('%d.%m.%Y')}[/bold yellow]")
            continue
        if date_obj > today:
            print("[bold red]Невозможно сформировать отчет на будущее[/bold red]")
            suggested = suggest_previous_valid_date(today, min_date, holidays_us)
            if suggested:
                print(f"[bold yellow]Предлагаемая дата: {suggested.strftime('%d.%m.%Y')}[/bold yellow]")
            continue
        # --- Конец нового блока ---

        if date_obj < min_date:
            print("[bold red]Ошибка: Указанный вами период не может быть применен, так как выходит за период деятельности N1 Broker[/bold red]")
            continue

        if start_date and date_obj <= start_date:
            print("[bold red]Ошибка: Неверный диапазон дат. Конечная дата должна быть старше начальной[/bold red]")
            continue

        if is_weekend(date_obj):
            print(f"[bold yellow]{date_obj.strftime('%d.%m.%Y')} — выходной день[/bold yellow]")
            before, after = find_nearest_valid_dates(date_obj, min_date, holidays_us)
            print(f"[bold yellow]Ближайшие доступные даты для формирования отчета: [/bold yellow]"
                  f"[bold cyan]{before.strftime('%d.%m.%Y')}[/bold cyan], [bold cyan]{after.strftime('%d.%m.%Y')}[/bold cyan]")
            continue

        if is_us_holiday(date_obj, holidays_us):
            holiday_name = holidays_us.get(date_obj)
            print(f"[bold yellow]{date_obj.strftime('%d.%m.%Y')} — {holiday_name}[/bold yellow]")
            before, after = find_nearest_valid_dates(date_obj, min_date, holidays_us)
            print(f"[bold yellow]Ближайшие доступные даты для формирования отчета: [/bold yellow]"
                  f"[bold cyan]{before.strftime('%d.%m.%Y')}[/bold cyan], [bold cyan]{after.strftime('%d.%m.%Y')}[/bold cyan]")
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
    print(f"[bold green]Даты сохранены в {path}[/bold green]")

def main():
    """
    Основная функция скрипта.
    """
    print_welcome()
    min_date = datetime.date(2022, 1, 1)
    holidays_us = holidays.US(years=range(2022, datetime.date.today().year + 2))

    # Ввод начальной даты
    start_date = get_date_input("Введите дату начала отчета (dd/mm/yyyy): ", min_date, holidays_us)
    print(f"[bold green]Дата начала отчета:[/bold green] [bold cyan]{start_date.strftime('%d.%m.%Y')}[/bold cyan]")
    wait_for_esc()

    # Ввод конечной даты
    end_date = get_date_input("Введите дату завершения отчета (dd/mm/yyyy): ", min_date, holidays_us, start_date=start_date)
    print(f"[bold green]Дата завершения отчета:[/bold green] [bold cyan]{end_date.strftime('%d.%m.%Y')}[/bold cyan]")
    wait_for_esc()

    # Сохраняем в JSON
    save_path = os.path.join("Data_work", "report_dates.json")
    save_dates_to_json(start_date, end_date, save_path)

    # Итоговое сообщение о периоде отчета
    print("[bold magenta]\nОтчет будет сформирован за период:[/bold magenta]")
    print(f"[bold magenta]с {start_date.strftime('%d.%m.%Y')} по {end_date.strftime('%d.%m.%Y')}[/bold magenta]")

if __name__ == "__main__":
    main()
