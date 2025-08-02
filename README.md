# 📊 Report_K

Проект для автоматизированной генерации Excel-отчётов для клиентов N1 Broker.

## 🧱 Структура проекта

```
Report/
├── Data_in/              # Входные Excel-файлы с портфелем
├── Data_work/            # Выходные JSON и Excel-файлы
├── scripts/
│   ├── BAT/              # *.bat-файлы запуска
│   │   ├── insert_date.bat
│   │   ├── name_clients.bat
│   │   ├── template_creator.bat
│   │   └── main.bat      # Запускает все модули последовательно
│   └── PS1/              # PowerShell-обёртки
│       ├── run_insert_date.ps1
│       ├── run_name_clients.ps1
│       ├── run_template_creator.ps1
│       └── run_main.ps1  # Запускает insert_date → name_clients → template_creator
├── insert_date.py        # Модуль 1: выбор и валидация даты отчёта
├── name_clients.py       # Модуль 2: извлечение имени клиента из Excel
├── template_creator.py   # Модуль 3: формирование Excel-отчёта
├── main.py               # Python-альтернатива для запуска всех модулей
├── README.md
└── CHANGELOG.md
```

## ⚙️ Алгоритм работы

1. **insert_date** — интерактивно запрашивает дату начала и окончания отчёта, сохраняет их в `Data_work/date_range.json`.
2. **name_clients** — извлекает имя клиента из входного файла Excel и сохраняет в `Data_work/name_clients.json`.
3. **template_creator** — создаёт Excel-отчёт в `Data_work/портфель_Фамилия_Дата.xlsx` на основе шаблона.

## 🚀 Запуск

Запуск через `main.bat`:

```bat
scripts\BAT\main.bat
```

Или через Python:

```bash
python main.py
```

## 🧩 Принцип Lego

Каждый модуль — самостоятельный блок. Проект расширяется добавлением новых "кубиков", которые также подключаются через `.bat` / `.ps1`.

---

### 🛠️ Требования

- Python 3.10
- PowerShell 7.5.2+
- Установленные библиотеки из `requirements.txt`
