# template_creator.py
"""
Скрипт для создания Excel-шаблона отчета с листом портфель.
Создает файл с именем клиента и диапазоном дат, архивирует старые шаблоны.
"""

import os
import json
import shutil
import datetime
from pathlib import Path

# Импорт xlwings для работы с Excel
try:
    import xlwings as xw
except ImportError:
    print("Устанавливаю xlwings для работы с Excel...")
    os.system(f'"{os.sys.executable}" -m pip install xlwings')
    import xlwings as xw


def load_json_data(path: str) -> dict:
    """
    Загружает данные из JSON-файла по указанному пути.
    
    Args:
        path (str): Путь к JSON-файлу
        
    Returns:
        dict: Загруженные данные из JSON-файла
        
    Raises:
        FileNotFoundError: Если файл не найден
        json.JSONDecodeError: Если файл содержит некорректный JSON
    """
    try:
        with open(path, 'r', encoding='utf-8') as file:
            return json.load(file)
    except FileNotFoundError:
        raise FileNotFoundError(f"Файл {path} не найден")
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError(f"Ошибка чтения JSON в файле {path}: {e}")


def get_output_filename(name_data: dict, date_data: dict) -> str:
    """
    Формирует имя файла в формате: портфель_Фамилия И. О._дата_дата.xlsx
    
    Args:
        name_data (dict): Данные с именем клиента из name_clients.json
        date_data (dict): Данные с датами из report_dates.json
        
    Returns:
        str: Сформированное имя файла
        
    Example:
        >>> name_data = {"surname": "Иванов", "initials": "И. В."}
        >>> date_data = {"start_date": "01.06.2024", "end_date": "30.06.2025"}
        >>> get_output_filename(name_data, date_data)
        'портфель_Иванов И. В._01.06.2024_30.06.2025.xlsx'
    """
    # Извлекаем данные из словарей
    #surname = name_data.get('surname', '')
    #initials = name_data.get('initials', '')
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
    
    # Формируем полное имя клиента
    full_name = f"{surname} {initials}".strip()
    
    # Создаем имя файла
    filename = f"портфель_{full_name}_{start_date}_{end_date}.xlsx"
    
    return filename


def archive_existing_portfolio_files(folder: str, backup_folder: str) -> list[str]:
    """
    Перемещает старые файлы портфеля в папку резервных копий.
    
    Args:
        folder (str): Папка для поиска файлов (Data_work)
        backup_folder (str): Папка для резервных копий (Data_Backup)
        
    Returns:
        list[str]: Список имен перемещенных файлов
        
    Note:
        Перемещает только файлы, начинающиеся с "портфель"
    """
    moved_files = []
    
    # Создаем папку для резервных копий, если её нет
    os.makedirs(backup_folder, exist_ok=True)
    
    # Ищем файлы, начинающиеся с "портфель"
    for filename in os.listdir(folder):
        if filename.startswith("портфель") and filename.endswith(".xlsx"):
            source_path = os.path.join(folder, filename)
            dest_path = os.path.join(backup_folder, filename)
            
            try:
                # Перемещаем файл
                shutil.move(source_path, dest_path)
                moved_files.append(filename)
                print(f"Найден файл с именем {filename}. Перемещен в папку Data_Backup")
            except Exception as e:
                print(f"Ошибка при перемещении файла {filename}: {e}")
    
    return moved_files


def create_excel_template(output_path: str, filename: str):
    """
    Создает Excel-файл с листом портфель и окрашивает ярлык в коричневый цвет.
    
    Args:
        output_path (str): Полный путь для сохранения файла
        filename (str): Имя создаваемого файла
        
    Note:
        Использует xlwings для создания Excel-файла
        Устанавливает коричневый цвет для ярлыка листа
    """
    try:
        # Создаем новую книгу
        app = xw.App(visible=False)
        wb = app.books.add()
        
        # Переименовываем первый лист в "портфель"
        sheet = wb.sheets[0]
        sheet.name = "портфель"
        
        # Устанавливаем коричневый цвет для ярлыка листа
        # RGB для коричневого: (139, 69, 19)
        # sheet.api.Tab.Color = (139, 69, 19)
        sheet.api.Tab.ColorIndex = 53  # Коричневый
        
        # Сохраняем файл
        wb.save(output_path)
        wb.close()
        app.quit()
        
    except Exception as e:
        print(f"Ошибка при создании Excel-файла: {e}")
        raise


def main():
    """
    Главная функция - организует весь процесс создания шаблона.
    
    Процесс:
    1. Загружает данные из JSON-файлов
    2. Формирует имя выходного файла
    3. Архивирует существующие файлы портфеля
    4. Создает новый Excel-шаблон
    5. Выводит финальное сообщение
    """
    # Определяем пути к файлам и папкам
    data_work_path = r"F:\Python Projets\Report\Data_work"
    data_backup_path = r"F:\Python Projets\Report\Data_Backup"
    
    name_clients_path = os.path.join(data_work_path, "name_clients.json")
    report_dates_path = os.path.join(data_work_path, "report_dates.json")
    
    try:
        # 1. Загружаем данные из JSON-файлов
        print("Загружаю данные из JSON-файлов...")
        name_data = load_json_data(name_clients_path)
        date_data = load_json_data(report_dates_path)
        
        # 2. Формируем имя выходного файла
        filename = get_output_filename(name_data, date_data)
        output_path = os.path.join(data_work_path, filename)
        
        print(f"Будет создан файл: {filename}")
        
        # 3. Архивируем существующие файлы портфеля
        print("Проверяю наличие старых файлов портфеля...")
        moved_files = archive_existing_portfolio_files(data_work_path, data_backup_path)
        
        if moved_files:
            print(f"Перемещено файлов: {len(moved_files)}")
        else:
            print("Старые файлы портфеля не найдены")
        
        # 4. Создаем новый Excel-шаблон
        print("Создаю Excel-шаблон...")
        create_excel_template(output_path, filename)
        
        # 5. Финальное сообщение
        print(f"\nФайл шаблона отчета «портфель» с именем {filename} создан.")
        print(f"Путь к файлу: {output_path}")
        
    except FileNotFoundError as e:
        print(f"Ошибка: {e}")
        print("Убедитесь, что файлы name_clients.json и report_dates.json существуют в папке Data_work")
    except json.JSONDecodeError as e:
        print(f"Ошибка чтения JSON: {e}")
        print("Проверьте корректность JSON-файлов")
    except Exception as e:
        print(f"Неожиданная ошибка: {e}")


if __name__ == "__main__":
    main()
