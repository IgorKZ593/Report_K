#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
map_instruments.py — Этап 1.
Каркас: поиск входного JSON isin_Фамилия И.О._start__end.json в Data_work,
чтение и валидация, печать статуса. Никакой логики справочников/записи на диск.
"""

import os
import sys
import json
import re
from glob import glob
from pathlib import Path
from typing import Tuple, List, Dict, Any, Optional

# === Автоустановка rich для цветного вывода ===
try:
    from rich.console import Console
    from rich import print
except ImportError:
    os.system(f'"{sys.executable}" -m pip install rich')
    from rich.console import Console
    from rich import print

console = Console()

# Константы путей (следуем принятой структуре проекта)
BASE_DIR = r"F:\Python Projets\Report"
DATA_WORK = BASE_DIR + r"\Data_work"

# ---------- Утилиты ----------

def parse_payload_name_from_filename(path: Path) -> Tuple[str, str, str]:
    """
    Из имени файла вида:
      isin_Фамилия И.О._DD.MM.YYYY__DD.MM.YYYY.json
    извлекает (client_str, start_date, end_date).
    Возвращает строки без изменений.
    """
    name = path.name
    # Жёсткое соответствие шаблону
    m = re.fullmatch(r"isin_(.+)_(\d{2}\.\d{2}\.\d{4})__(\d{2}\.\d{2}\.\d{4})\.json", name)
    if not m:
        raise ValueError(f"Имя файла не соответствует шаблону: {name}")
    client_str, start_date, end_date = m.group(1), m.group(2), m.group(3)
    return client_str, start_date, end_date


def load_client_isins(path: Path) -> Tuple[str, Dict[str, str], List[str]]:
    """
    Читает входной JSON и возвращает:
      client: строка из JSON (как есть)
      period: словарь {"start_date": "...", "end_date": "..."}
      isin_list: список ISIN
    Бросает осмысленные ошибки при проблемах с чтением/структурой.
    """
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        raise FileNotFoundError(f"Файл не найден: {path}")
    except json.JSONDecodeError as e:
        raise ValueError(f"Ошибка парсинга JSON в {path}: {e}")

    if not isinstance(data, dict):
        raise ValueError("Ожидался объект JSON верхнего уровня")

    client = data.get("client")
    period = data.get("period")
    isins = data.get("isin")

    if not isinstance(client, str) or not client.strip():
        raise ValueError("Поле 'client' отсутствует или пустое")

    if not isinstance(period, dict) or "start_date" not in period or "end_date" not in period:
        raise ValueError("Поле 'period' отсутствует или неполное")

    if not isinstance(isins, list) or not all(isinstance(x, str) for x in isins):
        raise ValueError("Поле 'isin' должно быть списком строк")

    return client, {"start_date": period["start_date"], "end_date": period["end_date"]}, isins


def find_input_payload(data_work: str) -> Path:
    """
    Ищет ровно один файл по маске isin_*.json в DATA_WORK.
    0 — ошибка; >1 — перечислить и ошибка; иначе вернуть Path.
    """
    pattern = os.path.join(data_work, "isin_*.json")
    files = [Path(p) for p in glob(pattern)]
    if not files:
        console.print(f"[red]❌ Во входной папке [/red][bright_cyan]{data_work}[/bright_cyan][red] нет файлов по маске isin_*.json[/red]")
        sys.exit(1)
    if len(files) > 1:
        console.print(f"[red]❌ Найдено несколько файлов по маске в [/red][bright_cyan]{data_work}[/bright_cyan][red]:[/red]")
        for p in files:
            console.print(f"[bright_cyan]  - {p.name}[/bright_cyan]")
        sys.exit(1)
    return files[0]

# ---------- Точка входа ----------

def main(argv: Optional[List[str]] = None) -> int:
    try:
        console.print("[bold green]🧭 map_instruments — Этап 1 (каркас)[/bold green]")
        console.print(f"[bright_cyan]Поиск входного файла в: {DATA_WORK}[/bright_cyan]")

        input_path = find_input_payload(DATA_WORK)

        # Имя файла → ожидаемые client/start/end (по имени)
        client_from_name, start_from_name, end_from_name = parse_payload_name_from_filename(input_path)

        console.print(f"[green]✅ Найден входной JSON: [/green][bright_cyan]{input_path.name}[/bright_cyan]")
        console.print(f"[green]↳ Ожидается из имени: client=[/green][bright_cyan]{client_from_name}[/bright_cyan][green], "
                      f"period=[/green][bright_cyan]{start_from_name}..{end_from_name}[/bright_cyan]")

        # Фактическое содержимое JSON
        client, period, isins = load_client_isins(input_path)
        console.print(f"[green]✅ Загружен JSON. Клиент:[/green] [bright_cyan]{client}[/bright_cyan]")
        console.print(f"[green]↳ Период:[/green] [bright_cyan]{period['start_date']}..{period['end_date']}[/bright_cyan]")
        console.print(f"[green]↳ Кол-во ISIN:[/green] [bright_cyan]{len(isins)}[/bright_cyan]")

        # Никаких сопоставлений/записей на диск на этом этапе
        console.print("[yellow]Этап 1 завершён: входной файл распознан и загружен. Дальнейшая логика будет добавлена на следующих этапах.[/yellow]")
        return 0

    except KeyboardInterrupt:
        console.print("\n[red]Операция прервана пользователем[/red]")
        return 1
    except Exception as e:
        console.print(f"[red]❌ Критическая ошибка: {e}[/red]")
        return 1


if __name__ == "__main__":
    sys.exit(main())
