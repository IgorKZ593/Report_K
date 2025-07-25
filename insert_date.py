import os
import sys
import json
import datetime
import holidays

# Импорт необходимых модулей и автоустановка недостающих
try:
    from prompt_toolkit import PromptSession
    from prompt_toolkit.key_binding import KeyBindings
    from prompt_toolkit.application import get_app_or_none
except ImportError:
    print("[bold yellow]Устанавливаю prompt_toolkit для интерактивного ввода...[/bold yellow]")
    os.system(f'"{sys.executable}" -m pip install prompt_toolkit')
    from prompt_toolkit import PromptSession
    from prompt_toolkit.key_binding import KeyBindings
    from prompt_toolkit.application import get_app_or_none

try:
    from rich import print
except ImportError:
    print("Устанавливаю rich для цветного вывода...")
    os.system(f'"{sys.executable}" -m pip install rich')
    from rich import print

# Проверка наличия необходимых внешних модулей
REQUIRED_MODULES = ["holidays", "rich"]
for mod in REQUIRED_MODULES:
    try:
        __import__(mod)
    except ImportError:
        print(f"[bold red]Модуль '{mod}' не установлен. Установите его: pip install {mod}[/bold red]")
        sys.exit(1)

# Удаление старого файла с датами, если он существует
json_path = os.path.join("Data_work", "report_dates.json")
if os.path.exists(json_path):
    try:
        os.remove(json_path)
    except Exception:
        pass

def print_welcome():
    print("[bold green]Подготовка аналитического отчета для клиентов N1 Broker[/bold green]")

def is_weekend(date_obj):
    return date_obj.weekday() >= 5

def is_us_holiday(date_obj, holidays_us):
    return date_obj in holidays_us

def find_nearest_valid_dates(date_obj, min_date, holidays_us):
    before = date_obj - datetime.timedelta(days=1)
    after = date_obj + datetime.timedelta(days=1)
    while before >= min_date:
        if not is_weekend(before) and not is_us_holiday(before, holidays_us):
            break
        before -= datetime.timedelta(days=1)
    while True:
        if not is_weekend(after) and not is_us_holiday(after, holidays_us):
            break
        after += datetime.timedelta(days=1)
    return before, after

def suggest_previous_valid_date(date_obj, min_date, holidays_us):
    prev_date = date_obj - datetime.timedelta(days=1)
    while prev_date >= min_date:
        if not is_weekend(prev_date) and not is_us_holiday(prev_date, holidays_us):
            return prev_date
        prev_date -= datetime.timedelta(days=1)
    return None

def get_date_input(prompt_text, min_date, holidays_us, start_date=None):
    #session = PromptSession()
    session = PromptSession(editing_mode="VI")
    bindings = KeyBindings()

    @bindings.add('escape')
    def _(event):
        app = get_app_or_none()
        if app:
            app.exit(result="__exit__")

    try:
        while True:
            result = session.prompt(prompt_text, key_bindings=bindings)
            if result == "__exit__":
                print("\n[bold red]Выход из программы по Esc[/bold red]")
                sys.exit(0)
            date_str = result.strip()

            try:
                date_obj = datetime.datetime.strptime(date_str, "%d/%m/%Y").date()
            except ValueError:
                print("[bold red]Ошибка: некорректный формат. Используйте dd/mm/yyyy[/bold red]")
                continue

            today = datetime.date.today()

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

            if date_obj < min_date:
                print("[bold red]Дата вне допустимого диапазона (до 2022 года)[/bold red]")
                continue

            if start_date and date_obj <= start_date:
                print("[bold red]Конечная дата должна быть позже начальной[/bold red]")
                continue

            if is_weekend(date_obj):
                print(f"[bold yellow]{date_obj.strftime('%d.%m.%Y')} — выходной день[/bold yellow]")
                before, after = find_nearest_valid_dates(date_obj, min_date, holidays_us)
                print(f"[bold yellow]Ближайшие допустимые даты: [/bold yellow]"
                      f"[bold cyan]{before.strftime('%d.%m.%Y')}[/bold cyan], "
                      f"[bold cyan]{after.strftime('%d.%m.%Y')}[/bold cyan]")
                continue

            if is_us_holiday(date_obj, holidays_us):
                holiday_name = holidays_us.get(date_obj)
                print(f"[bold yellow]{date_obj.strftime('%d.%m.%Y')} — {holiday_name}[/bold yellow]")
                before, after = find_nearest_valid_dates(date_obj, min_date, holidays_us)
                print(f"[bold yellow]Ближайшие допустимые даты: [/bold yellow]"
                      f"[bold cyan]{before.strftime('%d.%m.%Y')}[/bold cyan], "
                      f"[bold cyan]{after.strftime('%d.%m.%Y')}[/bold cyan]")
                continue

            return date_obj
    except (KeyboardInterrupt, EOFError):
        print("\n[bold red]Прерывание пользователем[/bold red]")
        sys.exit(0)

def save_dates_to_json(start_date, end_date, path):
    data = {
        "start_date": start_date.strftime("%d.%m.%Y"),
        "end_date": end_date.strftime("%d.%m.%Y")
    }
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[bold green]Даты сохранены в {path}[/bold green]")

def main():
    print_welcome()
    print("[bold yellow]Для выхода нажмите Ctrl+C или Esc дважды 🙂[/bold yellow]")
    min_date = datetime.date(2022, 1, 1)
    holidays_us = holidays.US(years=range(2022, datetime.date.today().year + 2))

    start_date = get_date_input("Введите дату начала отчета (dd/mm/yyyy): ", min_date, holidays_us)
    print(f"[bold green]Дата начала отчета: [bold cyan]{start_date.strftime('%d.%m.%Y')}[/bold cyan]")

    end_date = get_date_input("Введите дату завершения отчета (dd/mm/yyyy): ", min_date, holidays_us, start_date=start_date)
    print(f"[bold green]Дата завершения отчета: [bold cyan]{end_date.strftime('%d.%m.%Y')}[/bold cyan]")

    save_path = os.path.join("Data_work", "report_dates.json")
    save_dates_to_json(start_date, end_date, save_path)

    print("[bold magenta]\nОтчет будет сформирован за период:[/bold magenta]")
    print(f"[bold magenta]с {start_date.strftime('%d.%m.%Y')} по {end_date.strftime('%d.%m.%Y')}[/bold magenta]")

if __name__ == "__main__":
    main()

