# Report_K
Automated Report Generation in Python A structured project for financial report generation using Excel files, reference dictionaries (ISIN, tickers), and modular processing logic. Developed with versioning, reproducibility, and clarity in mind.

Проект предназначен для обработки и генерации отчетов на основе входных Excel-данных, справочников тикеров и ISIN.
📁 Структура проекта
Report/
├── .venv/             # Виртуальное окружение Python (на GitHub не сохраняется)
├── Data_in/           # Входные данные (исходные Excel-файлы, загружаемые пользователем)
├── Data_out/          # Финальные отчеты, готовые к передаче клиенту
├── Data_work/         # Временные файлы, промежуточные расчеты, логика обработки
├── dictionaries/      # Справочники ISIN, тикеров, валют и др.
└── main.py            # Главный скрипт запуска проекта
⚙️ Используемая версия Python
Python 3.10

Виртуальное окружение создается командой:
python -m venv .venv
🗂 Ветки проекта
- main — основная ветка
- Insert_date — разработка логики вставки даты и формирования отчетов
🚀 Запуск
1. Убедитесь, что активировано виртуальное окружение:
   .venv\Scripts\activate
2. Установите зависимости:
   pip install -r requirements.txt
3. Запустите проект:
   python main.py
