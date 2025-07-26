# Импорт необходимых модулей
import os
import sys
import json
import glob
from pathlib import Path


import os
import sys

# === Автоустановка rich (в первую очередь) ===
try:
    from rich import print
except ImportError:
    print("Устанавливаю rich для цветного вывода...")
    os.system(f'"{sys.executable}" -m pip install rich')
    from rich import print

# === Автоустановка xlwings ===
try:
    import xlwings as xw
except ImportError:
    print("[bold yellow]Устанавливаю xlwings...[/bold yellow]")
    os.system(f'"{sys.executable}" -m pip install xlwings')
    try:
        import xlwings as xw
    except ImportError:
        print("[bold red]Модуль xlwings не установлен. Установите вручную: pip install xlwings[/bold red]")
        sys.exit(1)


# Импорт xlwings для работы с Excel-файлами
try:
    import xlwings as xw
except ImportError:
    print("[bold red]Модуль xlwings не установлен. Установите его: pip install xlwings[/bold red]")
    sys.exit(1)

# Импорт rich для цветного вывода
try:
    from rich import print
except ImportError:
    print("Устанавливаю rich для цветного вывода...")
    os.system(f'"{sys.executable}" -m pip install rich')
    from rich import print

# Константы путей
DATA_IN_PATH = r"F:\Python Projets\Report\Data_in"
DATA_WORK_PATH = r"F:\Python Projets\Report\Data_work"
OUTPUT_FILE = os.path.join(DATA_WORK_PATH, "name_clients.json")

def find_report_files():
    """
    Ищет Excel-файлы с названием, содержащим слово 'отчет' (регистр не учитывается).
    
    Returns:
        list: Список найденных файлов
    """
    if not os.path.exists(DATA_IN_PATH):
        print(f"[bold red]Папка {DATA_IN_PATH} не найдена[/bold red]")
        return []
    
    # Поиск всех Excel-файлов с 'отчет' в названии (регистр не учитывается)
    pattern = os.path.join(DATA_IN_PATH, "*отчет*.xlsx")
    report_files = glob.glob(pattern, recursive=False)
    
    # Дополнительный поиск с другим регистром
    pattern2 = os.path.join(DATA_IN_PATH, "*ОТЧЕТ*.xlsx")
    report_files.extend(glob.glob(pattern2, recursive=False))
    
    # Удаляем дубликаты и сортируем
    report_files = list(set(report_files))
    report_files = [f for f in report_files if not os.path.basename(f).startswith('~$')]
    report_files.sort()
    return report_files
    


def validate_single_report_file(report_files):
    """
    Проверяет, что найден ровно один файл отчета.
    
    Args:
        report_files (list): Список найденных файлов отчетов
        
    Returns:
        str: Путь к единственному файлу отчета или None
    """
    if not report_files:
        print("[bold red]В папке Data_in не найдено файлов с названием, содержащим 'отчет'[/bold red]")
        return None
    
    if len(report_files) > 1:
        print("[bold red]Найдено несколько файлов с названием, содержащим 'отчет':[/bold red]")
        for file in report_files:
            print(f"  - {os.path.basename(file)}")
        print("[bold yellow]Оставьте только один файл отчета в папке Data_in[/bold yellow]")
        return None
    
    return report_files[0]

def check_portfolio_sheet(file_path):
    """
    Проверяет наличие листа 'портфель' в Excel-файле.
    
    Args:
        file_path (str): Путь к Excel-файлу
        
    Returns:
        bool: True если лист найден, False в противном случае
    """
    try:
        # Открываем файл с xlwings
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        
        # Проверяем наличие листа 'портфель' (регистр не учитывается)
        portfolio_sheet = None
        for sheet in wb.sheets:
            if sheet.name.lower() == 'портфель':
                portfolio_sheet = sheet
                break
        
        wb.close()
        app.quit()
        
        if portfolio_sheet is None:
            print("[bold red]В исходном файле отсутствует лист 'портфель'[/bold red]")
            print("[bold yellow]Проверьте и/или замените источник данных[/bold yellow]")
            return False
        
        return True
        
    except Exception as e:
        print(f"[bold red]Ошибка при открытии файла: {e}[/bold red]")
        return False

def extract_client_name(file_path):
    """
    Извлекает имя клиента из первой ячейки столбца 'Владелец счета'.
    
    Args:
        file_path (str): Путь к Excel-файлу
        
    Returns:
        str: Имя клиента или None в случае ошибки
    """
    try:
        # Открываем файл с xlwings
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        
        # Находим лист 'портфель'
        portfolio_sheet = None
        for sheet in wb.sheets:
            if sheet.name.lower() == 'портфель':
                portfolio_sheet = sheet
                break
        
        if portfolio_sheet is None:
            wb.close()
            app.quit()
            return None
        
        # Ищем столбец 'Владелец счета'
        used_range = portfolio_sheet.used_range
        header_row = used_range.rows[0]
        
        owner_column = None
        for cell in header_row:
            if cell.value and 'владелец счета' in str(cell.value).lower():
                owner_column = cell.column
                break
        
        if owner_column is None:
            print("[bold red]Столбец 'Владелец счета' не найден в листе 'портфель'[/bold red]")
            wb.close()
            app.quit()
            return None
        
        # Получаем значение из первой ячейки столбца (после заголовка)
        #client_name_cell = portfolio_sheet.range(f"{owner_column}2")
        #client_name = client_name_cell.value
        client_name_cell = portfolio_sheet.cells(2, owner_column)
        client_name = client_name_cell.value
        
        wb.close()
        app.quit()
        
        if not client_name:
            print("[bold red]Ячейка с именем клиента пуста[/bold red]")
            return None
        
        return str(client_name).strip()
        
    except Exception as e:
        print(f"[bold red]Ошибка при извлечении имени клиента: {e}[/bold red]")
        return None

def save_client_name_to_json(client_name):
    """
    Сохраняет имя клиента в JSON-файл.
    
    Args:
        client_name (str): Имя клиента для сохранения
        
    Returns:
        bool: True если сохранение успешно, False в противном случае
    """
    try:
        # Создаем папку Data_work, если она не существует
        os.makedirs(DATA_WORK_PATH, exist_ok=True)
        
        # Формируем данные для сохранения
        data = {
            "client_name": client_name
        }
        
        # Сохраняем в JSON с поддержкой UTF-8
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print(f"[bold green][✔] Имя клиента сохранено в {os.path.basename(OUTPUT_FILE)}[/bold green]")
        return True
        
    except Exception as e:
        print(f"[bold red]Ошибка при сохранении файла: {e}[/bold red]")
        return False

def get_user_confirmation():
    """
    Запрашивает подтверждение у пользователя.
    
    Returns:
        bool: True если пользователь подтвердил, False в противном случае
    """
    try:
        response = input("Сохранить имя клиента в файл name_clients.json? [Y/n]: ").strip().lower()
        return response in ['', 'y', 'yes', 'да', 'д']
    except (KeyboardInterrupt, EOFError):
        print("\n[bold red]Ввод прерван[/bold red]")
        return False

def main():
    """
    Основная функция модуля - выполняет весь процесс извлечения и сохранения имени клиента.
    """
    print("[bold green]Извлечение имени клиента из отчета[/bold green]")
    print(f"[bold yellow]Поиск файлов в: {DATA_IN_PATH}[/bold yellow]")
    
    # Шаг 1: Поиск файлов отчетов
    report_files = find_report_files()
    
    # Шаг 2: Проверка на единственность файла
    report_file = validate_single_report_file(report_files)
    if not report_file:
        sys.exit(1)
    
    print(f"[bold green]Найден файл: {os.path.basename(report_file)}[/bold green]")
    
    # Шаг 3: Проверка наличия листа 'портфель'
    if not check_portfolio_sheet(report_file):
        sys.exit(1)
    
    # Шаг 4: Извлечение имени клиента
    client_name = extract_client_name(report_file)
    if not client_name:
        sys.exit(1)
    
    print(f"[bold green]Обнаружено имя клиента: {client_name}[/bold green]")
    
    # Шаг 5: Запрос подтверждения
    if not get_user_confirmation():
        print(f"[bold yellow][!] Проверьте источник данных в папке {DATA_IN_PATH}[/bold yellow]")
        sys.exit(0)
    
    # Шаг 6: Сохранение в JSON
    if not save_client_name_to_json(client_name):
        sys.exit(1)
    
    print("[bold green]Обработка завершена успешно![/bold green]")

# Запуск модуля, если файл выполняется напрямую
if __name__ == "__main__":
    main()

