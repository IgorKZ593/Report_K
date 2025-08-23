import subprocess
import os

# Путь к .bat-файлам
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BAT_DIR = os.path.join(BASE_DIR, "scripts", "BAT")

# Список модулей в нужной последовательности
MODULES = [
    ("insert_date.bat", "📅 Ввод даты"),
    ("name_clients.bat", "👤 Имя клиента"),
    ("extract_isin.bat", "🔎 Извлечение ISIN"),
    ("template_creator.bat", "📄 Создание шаблона отчета")
]

def run_module(bat_file, description):
    print(f"\n[INFO] 🔸 Запуск модуля: {description}")
    path = os.path.join(BAT_DIR, bat_file)
    try:
        subprocess.run(path, check=True)
        print(f"[INFO] ✅ Завершено: {description}")
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] ❌ Ошибка при запуске {bat_file}: {e}")
        exit(1)

def main():
    print("=== 🚀 Запуск подготовки отчета N1 Broker ===")
    for bat, desc in MODULES:
        run_module(bat, desc)
    print("\n=== 🏁 Все этапы завершены успешно ===")

if __name__ == "__main__":
    main()
