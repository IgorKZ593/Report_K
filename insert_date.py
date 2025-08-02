# Импорт стандартных и внешних модулей
import holidays
import os
import sys
import json
import datetime

# Импорт prompt_toolkit для интерактивного ввода дат
try:
    from prompt_toolkit import PromptSession
except ImportError:
    # Если prompt_toolkit не установлен, устанавливаем его автоматически
    print("Устанавливаю prompt_toolkit для интерактивного ввода...")
    os.system(f'"{sys.executable}" -m pip install prompt_toolkit')
    from prompt_toolkit import PromptSession

# Импорт rich для цветного вывода в консоль
try:
    from rich import print
except ImportError:
    # Если rich не установлен, устанавливаем его автоматически
    print("Устанавливаю rich для цветного вывода...")
    os.system(f'"{sys.executable}" -m pip install rich')
    from rich import print

# Проверка наличия необходимых внешних модулей (holidays, rich)
REQUIRED_MODULES = ["holidays", "rich"]
for mod in REQUIRED_MODULES:
    try:
        __import__(mod)
    except ImportError:
        print(f"[bold red]Модуль '{mod}' не установлен. Установите его: pip install {mod}[/bold red]")
        sys.exit(1)

# Удаление старого файла с датами, если он существует, чтобы избежать конфликтов при повторном запуске
json_path = os.path.join("Data_work", "report_dates.json")
if os.path.exists(json_path):
    try:
        os.remove(json_path)
    except Exception:
        pass

# Функция приветствия пользователя
# Выводит информационное сообщение о запуске скрипта

def print_welcome():
    print("[bold green]Подготовка аналитического отчета для клиентов N1 Broker[/bold green]")

# Проверка, является ли дата выходным днем (суббота или воскресенье)
def is_weekend(date_obj):
    return date_obj.weekday() >= 5

# Проверка, является ли дата праздничным днем в США (используется holidays)
def is_us_holiday(date_obj, holidays_us):
    return date_obj in holidays_us

# Поиск ближайших допустимых дат до и после заданной даты
# Исключаются выходные и праздничные дни
def find_nearest_valid_dates(date_obj, min_date, holidays_us):
    before = date_obj - datetime.timedelta(days=1)
    after = date_obj + datetime.timedelta(days=1)
    # Поиск предыдущей допустимой даты
    while before >= min_date:
        if not is_weekend(before) and not is_us_holiday(before, holidays_us):
            break
        before -= datetime.timedelta(days=1)
    # Поиск следующей допустимой даты
    while True:
        if not is_weekend(after) and not is_us_holiday(after, holidays_us):
            break
        after += datetime.timedelta(days=1)
    return before, after

# Предложение предыдущей допустимой даты, если текущая недопустима (например, выходной или праздник)
def suggest_previous_valid_date(date_obj, min_date, holidays_us):
    prev_date = date_obj - datetime.timedelta(days=1)
    while prev_date >= min_date:
        if not is_weekend(prev_date) and not is_us_holiday(prev_date, holidays_us):
            return prev_date
        prev_date -= datetime.timedelta(days=1)
    return None

# Основная функция для интерактивного ввода даты
# Использует prompt_toolkit для красивого и удобного ввода
# Вся логика выхода по Esc удалена, остался только Ctrl+C
# Валидация даты: формат, диапазон, выходные, праздники, последовательность

def get_date_input(prompt_text, min_date, holidays_us, start_date=None):
    """
    Запрашивает у пользователя ввод даты с помощью prompt_toolkit.
    Для выхода используйте Ctrl+C.
    Проверяет корректность формата, диапазона, выходных и праздничных дней.
    """
    session = PromptSession()
    try:
        while True:
            # Получаем строку от пользователя
            result = session.prompt(prompt_text)
            # Если пользователь прервал ввод (например, EOF), корректно завершаем
            if result is None:
                print("Ввод прерван")
                sys.exit(0)
            date_str = result.strip()
            # Проверка формата даты
            try:
                date_obj = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
            except ValueError:
                print("[bold red]Ошибка: некорректный формат. Используйте dd/mm/yyyy[/bold red]")
                continue

            today = datetime.date.today()

            # Проверка: нельзя выбрать сегодняшнюю дату
            if date_obj == today:
                print("[bold red]Невозможно сделать отчет на текущую дату[/bold red]")
                suggested = suggest_previous_valid_date(date_obj, min_date, holidays_us)
                if suggested:
                    print(f"[bold yellow]Предлагаемая дата: {suggested.strftime('%d.%m.%Y')}[/bold yellow]")
                continue

            # Проверка: нельзя выбрать будущую дату
            if date_obj > today:
                print("[bold red]Невозможно сформировать отчет на будущее[/bold red]")
                suggested = suggest_previous_valid_date(today, min_date, holidays_us)
                if suggested:
                    print(f"[bold yellow]Предлагаемая дата: {suggested.strftime('%d.%m.%Y')}[/bold yellow]")
                continue

            # Проверка: дата не должна быть меньше минимально допустимой
            if date_obj < min_date:
                print("[bold red]Дата вне допустимого диапазона (до 2022 года)[/bold red]")
                continue

            # Проверка: конечная дата должна быть позже начальной
            if start_date and date_obj <= start_date:
                print("[bold red]Конечная дата должна быть позже начальной[/bold red]")
                continue

            # Проверка: дата не должна быть выходным
            if is_weekend(date_obj):
                print(
                    f"[bold yellow]{date_obj.strftime('%d.%m.%Y')}[/bold yellow] — [bold red]выходной день[/bold red]")
                before, after = find_nearest_valid_dates(date_obj, min_date, holidays_us)
                print(f"[bold yellow]Ближайшие допустимые даты: [/bold yellow]"
                      f"[bold cyan]{before.strftime('%d.%m.%Y')}[/bold cyan], "
                      f"[bold cyan]{after.strftime('%d.%m.%Y')}[/bold cyan]")
                continue

            # Проверка: дата не должна быть праздничным днем
            if is_us_holiday(date_obj, holidays_us):
                holiday_name = holidays_us.get(date_obj)
                print(
                    f"[bold yellow]{date_obj.strftime('%d.%m.%Y')}[/bold yellow] — [bold red]{holiday_name}[/bold red]")
                before, after = find_nearest_valid_dates(date_obj, min_date, holidays_us)
                print(f"[bold yellow]Ближайшие допустимые даты: [/bold yellow]"
                      f"[bold cyan]{before.strftime('%d.%m.%Y')}[/bold cyan], "
                      f"[bold cyan]{after.strftime('%d.%m.%Y')}[/bold cyan]")
                continue

            # Если все проверки пройдены — возвращаем объект даты
            return date_obj
    except (KeyboardInterrupt, EOFError):
        # Обработка выхода по Ctrl+C или EOF
        print("\n[bold red]Выход из программы (Ctrl+C)[/bold red]")
        sys.exit(0)

# Сохраняет выбранные пользователем даты в JSON-файл для последующего использования

def save_dates_to_json(start_date, end_date, path):
    data = {
        "start_date": start_date.strftime("%d.%m.%Y"),
        "end_date": end_date.strftime("%d.%m.%Y")
    }
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[bold green]Даты сохранены в {path}[/bold green]")

# Главная функция — точка входа в программу
# Выводит приветствие, инструкции, запускает ввод дат, сохраняет результат и выводит итоговый диапазон

def main():
    print_welcome()
    print("[bold yellow]Для выхода нажмите Ctrl+C в любой момент[/bold yellow]")
    min_date = datetime.date(2022, 1, 1)
    holidays_us = holidays.US(years=range(2022, datetime.date.today().year + 2))

    # Ввод даты начала отчета
    start_date = get_date_input("Введите дату начала отчета (dd/mm/yyyy): ", min_date, holidays_us)
    print(f"[bold green]Дата начала отчета: [bold cyan]{start_date.strftime('%d.%m.%Y')}[/bold cyan]")

    # Ввод даты завершения отчета
    end_date = get_date_input("Введите дату завершения отчета (dd/mm/yyyy): ", min_date, holidays_us, start_date=start_date)
    print(f"[bold green]Дата завершения отчета: [bold cyan]{end_date.strftime('%d.%m.%Y')}[/bold cyan]")

    # Сохраняем выбранные даты в файл
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    save_path = os.path.join(BASE_DIR, "Data_work", "report_dates.json")

    save_dates_to_json(start_date, end_date, save_path)

    # Финальный вывод периода отчета
    print("[bold magenta]\nОтчет будет сформирован за период:[/bold magenta]")
    print(f"[bold magenta]с {start_date.strftime('%d.%m.%Y')} по {end_date.strftime('%d.%m.%Y')}[/bold magenta]")

# Запуск main(), если скрипт запущен напрямую
if __name__ == "__main__":
    main()

