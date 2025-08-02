# template_creator.py
"""
Скрипт для создания Excel-шаблона отчета с листом «портфель».
Создает файл с именем клиента и диапазоном дат, архивирует старые шаблоны.
"""

import os
import json
import shutil
from pathlib import Path

# ===============================
# 📦 Проверка и установка rich
# ===============================
try:
    from rich import print
    from rich.console import Console
except ImportError:
    import os
    print("📦 Устанавливаю библиотеку rich...")
    os.system(f'"{os.sys.executable}" -m pip install rich')
    from rich import print
    from rich.console import Console

console = Console()

# ===============================
# 📦 Проверка и установка xlwings
# ===============================
try:
    import xlwings as xw
except ImportError:
    console.print("📦 Устанавливаю библиотеку xlwings...", style="bold green")
    os.system(f'"{os.sys.executable}" -m pip install xlwings')
    import xlwings as xw


def load_json_data(path: str) -> dict:
    """Загружает данные из JSON-файла по указанному пути."""
    try:
        with open(path, 'r', encoding='utf-8') as file:
            return json.load(file)
    except FileNotFoundError:
        raise FileNotFoundError(f"Файл {path} не найден")
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError(f"Ошибка чтения JSON в файле {path}: {e}")


def get_output_filename(name_data: dict, date_data: dict) -> str:
    """Формирует имя файла: портфель_Фамилия И. О._дата_дата.xlsx"""
    client_name = name_data.get("client_name", "").strip()
    if client_name:
        parts = client_name.split(" ", 1)
        surname = parts[0] if len(parts) > 0 else ""
        initials = parts[1] if len(parts) > 1 else ""
    else:
        surname = ""
        initials = ""

    start_date = date_data.get('start_date', '')
    end_date = date_data.get('end_date', '')

    full_name = f"{surname} {initials}".strip()
    filename = f"портфель_{full_name}_{start_date}_{end_date}.xlsx"

    return filename


def archive_existing_portfolio_files(folder: str, backup_folder: str) -> list[str]:
    """Перемещает старые файлы портфеля в папку резервных копий."""
    moved_files = []
    os.makedirs(backup_folder, exist_ok=True)

    for filename in os.listdir(folder):
        if filename.startswith("портфель") and filename.endswith(".xlsx"):
            source_path = os.path.join(folder, filename)
            dest_path = os.path.join(backup_folder, filename)
            try:
                shutil.move(source_path, dest_path)
                moved_files.append(filename)
                console.print(f"📦 Найден файл [white]{filename}[/] → перемещён в [bold]Data_Backup[/]")
            except Exception as e:
                console.print(f"[red]Ошибка при перемещении файла {filename}:[/] {e}")

    return moved_files


def create_excel_template(output_path: str, filename: str):
    """Создает Excel-файл с листом «портфель» и коричневым ярлыком."""
    try:
        app = xw.App(visible=False)
        wb = app.books.add()

        sheet = wb.sheets[0]
        sheet.name = "портфель"
        sheet.api.Tab.ColorIndex = 53  # Коричневый

        wb.save(output_path)
        wb.close()
        app.quit()

    except Exception as e:
        console.print(f"[red]Ошибка при создании Excel-файла:[/] {e}")
        raise


def main():
    """
    Главная функция — организует весь процесс создания шаблона.
    """
    data_work_path = r"F:\Python Projets\Report\Data_work"
    data_backup_path = r"F:\Python Projets\Report\Data_Backup"

    name_clients_path = os.path.join(data_work_path, "name_clients.json")
    report_dates_path = os.path.join(data_work_path, "report_dates.json")

    try:
        # 1. Загружаем данные из JSON-файлов
        console.print(f"[bold cyan]📄 Загружаю данные из JSON-файлов...[/]")
        name_data = load_json_data(name_clients_path)
        date_data = load_json_data(report_dates_path)

        # 2. Формируем имя выходного файла (output_path)
        filename = get_output_filename(name_data, date_data)
        output_path = os.path.join(data_work_path, filename)

        console.print(f"[green]📁 Будет создан файл:[/] [white]{filename}[/]")

        # 3. Архивируем старые файлы портфеля
        console.print("[yellow]📦 Проверяю наличие старых файлов портфеля...[/]")
        moved_files = archive_existing_portfolio_files(data_work_path, data_backup_path)

        if moved_files:
            console.print(f"[magenta]🔁 Перемещено файлов:[/] {len(moved_files)}")
        else:
            console.print("[grey]⏳ Старые файлы портфеля не найдены[/]")

        # 4. Создаём новый Excel-шаблон
        console.print("[blue]🛠 Создаю Excel-шаблон...[/]")
        create_excel_template(output_path, filename)

        console.print(f"[bold green]✔️ Файл шаблона отчета создан:[/] [white]{filename}[/]")
        #console.print(f"[dim]📍 Путь к файлу:[/] {output_path}")
        console.print(f"[white]📍 Путь к файлу:[/] [bold cyan]{output_path}[/]")

    except FileNotFoundError as e:
        console.print(f"[red]❌ Ошибка: {e}[/]")
        console.print("[yellow]⚠️ Убедитесь, что файлы name_clients.json и report_dates.json существуют в папке Data_work[/]")
    except json.JSONDecodeError as e:
        console.print(f"[red]❌ Ошибка чтения JSON: {e}[/]")
        console.print("[yellow]⚠️ Проверьте корректность JSON-файлов[/]")
    except Exception as e:
        console.print(f"[bold red]💥 Неожиданная ошибка:[/] {e}")


if __name__ == "__main__":
    main()

