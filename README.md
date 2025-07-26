# 📊 Проект: Автоматизация клиентского отчета

Этот проект предназначен для автоматизированной подготовки Excel-отчета по инвестиционному портфелю клиента. Структура отчета и его логика строго регламентированы и разделены на независимые блоки с собственными скриптами запуска.

---

## 📁 Структура проекта

```
Report/
├── Data_work/              # Рабочие файлы (JSON, Excel)
│   └── .gitkeep
├── Data_in/                # Входные данные
│   └── .gitkeep
├── Data_out/               # Финальные выгрузки (по желанию)
│   └── .gitkeep
├── Data_Backup/            # Резервные копии ранее созданных файлов
│   └── .gitkeep
├── logs/                   # Логи работы (опционально)
│   └── .gitkeep
├── main.py                 # Основной управляющий скрипт (если используется)
├── insert_date.py          # Блок для получения дат отчета
├── name_clients.py         # Блок для получения имени клиента
├── template_creator.py     # Блок создания шаблона Excel-отчета
├── run_insert_date.ps1     # PowerShell-запуск блока insert_date
├── run_name_clients.ps1    # PowerShell-запуск блока name_clients
├── run_template_creator.ps1# PowerShell-запуск шаблона отчета
├── .gitignore              # Git-исключения
└── README.md               # Документация
```

---

## ⚙️ Назначение ключевых модулей

| Файл                   | Назначение |
|------------------------|------------|
| `insert_date.py`       | Получение начальной и конечной даты отчета, сохранение в JSON |
| `name_clients.py`      | Извлечение имени клиента (Фамилия И. О.) и сохранение в JSON |
| `template_creator.py`  | Генерация пустого Excel-файла с именем клиента и датами |
| `main.py`              | Возможный управляющий скрипт (может объединять шаги) |

---

## 🚀 Как запускать

### Через PowerShell
```powershell
.
un_insert_date.ps1
.
un_name_clients.ps1
.
un_template_creator.ps1
```

### Или напрямую через Python (если активировано .venv)
```bash
python insert_date.py
python name_clients.py
python template_creator.py
```

---

## 📌 Требования

- Python 3.10+
- Установленные зависимости:
  ```bash
  pip install xlwings
  ```

---

---

© 2025 N1 Broker · Автор: Игорь
