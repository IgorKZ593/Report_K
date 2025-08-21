# template_creator.py
"""
Скрипт для создания Excel-шаблона отчета с листом «портфель».
Создает файл с именем клиента и диапазоном дат, архивирует старые шаблоны.

Основные функции:
- Загрузка данных из JSON-файлов (имя клиента и даты отчета)
- Формирование имени выходного файла
- Архивирование старых файлов портфеля
- Создание Excel-шаблона с двумя листами
"""

# ===============================
# Импорт стандартных библиотек
# ===============================
import os          # Для работы с файловой системой и путями
import sys         # Для доступа к sys.executable (путь к Python)
import json        # Для работы с JSON-файлами
import shutil      # Для перемещения файлов (архивирование)
from pathlib import Path  # Для работы с путями (альтернатива os.path)

# ===============================
# 📦 Проверка и установка rich
# ===============================
try:
    from rich import print
    from rich.console import Console
except ImportError:
    # Если rich не установлен, устанавливаем его автоматически
    print("📦 Устанавливаю библиотеку rich...")
    os.system(f'"{sys.executable}" -m pip install rich')
    from rich import print
    from rich.console import Console

# Создаем объект консоли для красивого вывода
console = Console()

# ===============================
# 📦 Проверка и установка xlwings
# ===============================
try:
    import xlwings as xw
except ImportError:
    # Если xlwings не установлен, устанавливаем его автоматически
    console.print("📦 Устанавливаю библиотеку xlwings...", style="bold green")
    os.system(f'"{sys.executable}" -m pip install xlwings')
    import xlwings as xw


def load_json_data(path: str) -> dict:
    """
    Загружает данные из JSON-файла по указанному пути.
    
    Параметры:
        path (str): Путь к JSON-файлу для загрузки
        
    Возвращает:
        dict: Словарь с данными из JSON-файла
        
    Исключения:
        FileNotFoundError: Если файл не найден по указанному пути
        json.JSONDecodeError: Если файл содержит некорректный JSON
        
    Пример использования:
        name_data = load_json_data("Data_work/name_clients.json")
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
    Формирует имя файла Excel в формате: портфель_Фамилия И. О._дата_дата.xlsx
    
    Параметры:
        name_data (dict): Словарь с данными клиента, должен содержать ключ "client_name"
        date_data (dict): Словарь с датами, должен содержать ключи "start_date" и "end_date"
        
    Возвращает:
        str: Имя файла в нужном формате
        
    Логика работы:
        1. Извлекает имя клиента из name_data
        2. Разделяет полное имя на фамилию и инициалы
        3. Извлекает даты начала и окончания из date_data
        4. Формирует имя файла по шаблону
        
    Пример:
        name_data = {"client_name": "Иванов Иван Петрович"}
        date_data = {"start_date": "01.01.2024", "end_date": "31.01.2024"}
        Результат: "портфель_Иванов И. П._01.01.2024_31.01.2024.xlsx"
    """
    # Извлекаем имя клиента из словаря
    client_name = name_data.get("client_name", "").strip()
    
    # Разделяем полное имя на фамилию и инициалы
    if client_name:
        parts = client_name.split(" ", 1)  # Разделяем по первому пробелу
        surname = parts[0] if len(parts) > 0 else ""      # Фамилия - первая часть
        initials = parts[1] if len(parts) > 1 else ""     # Инициалы - вторая часть
    else:
        surname = ""
        initials = ""

    # Извлекаем даты из словаря
    start_date = date_data.get('start_date', '')
    end_date = date_data.get('end_date', '')

    # Формируем полное имя и имя файла
    full_name = f"{surname} {initials}".strip()
    filename = f"портфель_{full_name}_{start_date}_{end_date}.xlsx"

    return filename


def archive_existing_portfolio_files(folder: str, backup_folder: str) -> list[str]:
    """
    Перемещает старые файлы портфеля в папку резервных копий.
    
    Параметры:
        folder (str): Папка, где искать файлы для архивирования
        backup_folder (str): Папка для сохранения резервных копий
        
    Возвращает:
        list[str]: Список имен перемещенных файлов
        
    Логика работы:
        1. Создает папку для резервных копий, если она не существует
        2. Просматривает все файлы в исходной папке
        3. Находит файлы, начинающиеся с "портфель" и заканчивающиеся на ".xlsx"
        4. Перемещает найденные файлы в папку резервных копий
        5. Выводит информацию о каждом перемещенном файле
        6. Возвращает список имен перемещенных файлов
        
    Пример использования:
        moved_files = archive_existing_portfolio_files("Data_work", "Data_Backup")
    """
    moved_files = []
    
    # Создаем папку для резервных копий, если она не существует
    os.makedirs(backup_folder, exist_ok=True)

    # Просматриваем все файлы в исходной папке
    for filename in os.listdir(folder):
        # Проверяем, подходит ли файл под критерии архивирования
        if filename.startswith("портфель") and filename.endswith(".xlsx"):
            source_path = os.path.join(folder, filename)
            dest_path = os.path.join(backup_folder, filename)
            
            try:
                # Перемещаем файл в папку резервных копий
                shutil.move(source_path, dest_path)
                moved_files.append(filename)
                console.print(f"📦 Найден файл [white]{filename}[/] → перемещён в [bold]Data_Backup[/]")
            except Exception as e:
                console.print(f"[red]Ошибка при перемещении файла {filename}:[/] {e}")

    return moved_files


def create_excel_template(output_path: str, filename: str):
    """
    Создает Excel-файл с двумя листами: "портфель" и "stock_etf_price".
    
    Параметры:
        output_path (str): Полный путь к создаваемому Excel-файлу
        filename (str): Имя файла (используется для логирования)
        
    Логика работы:
        1. Создает новый экземпляр Excel через xlwings
        2. Создает лист "портфель" с коричневой вкладкой
        3. Добавляет лист "stock_etf_price" с синей вкладкой
        4. Заполняет заголовки таблицы с форматированием
        5. Настраивает ширину столбцов
        6. Сохраняет файл и закрывает Excel
        
    Обработка ошибок:
        - Все операции форматирования обернуты в try-except
        - При ошибке корректно закрывает Excel и освобождает ресурсы
        
    Пример использования:
        create_excel_template("Data_work/portfolio.xlsx", "portfolio.xlsx")
    """
    app = None
    try:
        # Создаем новый экземпляр Excel (невидимый)
        app = xw.App(visible=False)
        wb = app.books.add()

        # ===============================
        # Создание листа "портфель"
        # ===============================
        sheet = wb.sheets[0]  # Получаем первый лист
        sheet.name = "портфель"  # Переименовываем его
        
        # Устанавливаем цвет вкладки (коричневый)
        try:
            sheet.api.Tab.ColorIndex = 53  # Коричневый цвет
        except:
            pass  # Если не удается установить цвет, продолжаем работу

        # ===============================
        # Создание листа "stock_etf_price"
        # ===============================
        stock_sheet = wb.sheets.add("stock_etf_price")  # Добавляем новый лист
        
        # Устанавливаем цвет вкладки (синий)
        try:
            stock_sheet.api.Tab.ColorIndex = 5  # Синий цвет
        except:
            pass

        # ===============================
        # Заполнение заголовков таблицы
        # ===============================
        # Определяем заголовки для листа stock_etf_price
        headers = ["ISIN", "Тикер", "Название", "start_date", "start_price", "end_date", "end_price", "Отклонение"]
        
        # Записываем заголовки в первую строку с форматированием
        for col, header in enumerate(headers, 1):
            cell = stock_sheet.range((1, col))  # Получаем ячейку в первой строке
            cell.value = header  # Устанавливаем значение
            
            # Применяем форматирование: жирный шрифт и выравнивание по центру
            try:
                cell.api.Font.Bold = True  # Жирный шрифт
                cell.api.HorizontalAlignment = -4108  # xlCenter - выравнивание по центру
                cell.api.VerticalAlignment = -4108    # xlCenter - выравнивание по центру
            except:
                pass  # Если форматирование не удается, продолжаем

        # ===============================
        # Настройка ширины столбцов
        # ===============================
        try:
            # Устанавливаем одинаковую ширину для всех столбцов (A-H)
            stock_sheet.api.Columns("A:H").ColumnWidth = 12
        except:
            pass

        # ===============================
        # Сохранение и закрытие
        # ===============================
        wb.save(output_path)  # Сохраняем файл
        wb.close()  # Закрываем книгу
        
        # Корректно закрываем Excel
        if app:
            app.quit()

    except Exception as e:
        # Обработка ошибок при создании файла
        console.print(f"[red]Ошибка при создании Excel-файла:[/] {e}")
        
        # Корректно закрываем Excel даже при ошибке
        if app:
            try:
                app.quit()
            except:
                pass
        raise  # Перебрасываем исключение дальше


def main():
    """
    Главная функция — организует весь процесс создания шаблона.
    
    Последовательность выполнения:
        1. Загружает данные из JSON-файлов
        2. Формирует имя выходного файла
        3. Архивирует старые файлы портфеля
        4. Создает новый Excel-шаблон
        5. Выводит информацию о результатах
        
    Обработка ошибок:
        - FileNotFoundError: Если не найдены JSON-файлы
        - json.JSONDecodeError: Если JSON-файлы повреждены
        - Exception: Для всех остальных ошибок
        
    Пути к файлам:
        - name_clients.json: содержит имя клиента
        - report_dates.json: содержит даты отчета
        - Выходной файл: создается в Data_work
        - Резервные копии: сохраняются в Data_Backup
    """
    # ===============================
    # Определение путей к файлам
    # ===============================
    data_work_path = r"F:\Python Projets\Report\Data_work"      # Папка с данными
    data_backup_path = r"F:\Python Projets\Report\Data_Backup"  # Папка для резервных копий

    # Формируем полные пути к JSON-файлам
    name_clients_path = os.path.join(data_work_path, "name_clients.json")
    report_dates_path = os.path.join(data_work_path, "report_dates.json")

    try:
        # ===============================
        # 1. Загружаем данные из JSON-файлов
        # ===============================
        console.print(f"[bold cyan]📄 Загружаю данные из JSON-файлов...[/]")
        name_data = load_json_data(name_clients_path)  # Загружаем данные клиента
        date_data = load_json_data(report_dates_path)  # Загружаем данные дат

        # ===============================
        # 2. Формируем имя выходного файла
        # ===============================
        filename = get_output_filename(name_data, date_data)  # Получаем имя файла
        output_path = os.path.join(data_work_path, filename)  # Формируем полный путь

        console.print(f"[green]📁 Будет создан файл:[/] [white]{filename}[/]")

        # ===============================
        # 3. Архивируем старые файлы портфеля
        # ===============================
        console.print("[yellow]📦 Проверяю наличие старых файлов портфеля...[/]")
        moved_files = archive_existing_portfolio_files(data_work_path, data_backup_path)

        # Выводим информацию о результатах архивирования
        if moved_files:
            console.print(f"[magenta]🔁 Перемещено файлов:[/] {len(moved_files)}")
        else:
            console.print("[grey]⏳ Старые файлы портфеля не найдены[/]")

        # ===============================
        # 4. Создаём новый Excel-шаблон
        # ===============================
        console.print("[blue]🛠 Создаю Excel-шаблон...[/]")
        create_excel_template(output_path, filename)

        # ===============================
        # 5. Выводим информацию об успешном создании
        # ===============================
        console.print(f"[bold green]✔️ Файл шаблона отчета создан:[/] [white]{filename}[/]")
        console.print(f"[white]📍 Путь к файлу:[/] [bold cyan]{output_path}[/]")

    except FileNotFoundError as e:
        # Обработка ошибки: файлы не найдены
        console.print(f"[red]❌ Ошибка: {e}[/]")
        console.print(
            "[yellow]⚠️ Убедитесь, что файлы name_clients.json и report_dates.json существуют в папке Data_work[/]")
    except json.JSONDecodeError as e:
        # Обработка ошибки: поврежденный JSON
        console.print(f"[red]❌ Ошибка чтения JSON: {e}[/]")
        console.print("[yellow]⚠️ Проверьте корректность JSON-файлов[/]")
    except Exception as e:
        # Обработка всех остальных ошибок
        console.print(f"[bold red]💥 Неожиданная ошибка:[/] {e}")


# ===============================
# Точка входа в программу
# ===============================
if __name__ == "__main__":
    # Запускаем главную функцию только если скрипт запущен напрямую
    # (не импортирован как модуль)
    main()
