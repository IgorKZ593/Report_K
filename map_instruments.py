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
    from rich.table import Table
except ImportError:
    os.system(f'"{sys.executable}" -m pip install rich')
    from rich.console import Console
    from rich import print
    from rich.table import Table

console = Console()

# Константы путей (следуем принятой структуре проекта)
BASE_DIR = r"F:\Python Projets\Report"
DATA_WORK = BASE_DIR + r"\Data_work"

# Пути к справочникам
REF_STOCKS_XLSX = r"F:\Python Projets\Report\dictionaries\reference_stocks\reference_stocks_etf.xlsx"
REF_BONDS_XLSX  = r"F:\Python Projets\Report\dictionaries\reference_bonds\reference_bonds.xlsx"
REF_SP_XLSX     = r"F:\Python Projets\Report\dictionaries\reference_structured\TS\TS.xlsx"
REF_SP_PDF_DIR  = r"F:\Python Projets\Report\dictionaries\reference_structured\TS"

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

# === Автоустановка openpyxl ===
try:
    from openpyxl import load_workbook
except ImportError:
    os.system(f'"{sys.executable}" -m pip install openpyxl')
    from openpyxl import load_workbook


def _norm_isin(s: str) -> str:
    return (s or "").strip().upper()


def load_reference_stocks(xlsx_path: str) -> dict:
    """
    Лист: 'акции_etf'
    Колонки: A=ISIN, B=Тикер, C=Название, D=Тип
    Возврат: { ISIN: {"ticker": str, "type": str, "name": str} }
    """
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    if "акции_etф" in wb.sheetnames:
        ws = wb["акции_etф"]
    else:
        ws = wb["акции_etf"]  # на случай опечаток регистра/раскладки
    ref = {}
    # пропускаем заголовок (первая строка)
    for row in ws.iter_rows(min_row=2, values_only=True):
        isin, ticker, name, typ = row[:4]
        isin = _norm_isin(isin)
        if not isin:
            continue
        ref[isin] = {
            "ticker": (ticker or "").strip(),
            "type": (typ or "").strip(),
            "name": (name or "").strip(),
        }
    return ref


def load_reference_bonds(xlsx_path: str) -> dict:
    """
    Лист: 'bonds'
    Колонки: A=ISIN, B=Название инструмента
    Возврат: { ISIN: {"name": str} }
    """
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["bonds"]
    ref = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        isin, name = row[:2]
        isin = _norm_isin(isin)
        if not isin:
            continue
        ref[isin] = {"name": (name or "").strip()}
    return ref


def load_reference_structured(xlsx_path: str, pdf_dir: str) -> dict:
    """
    Лист: 'TS'
    Колонки: B=ISIN, C=ссылка (необязательна для нас)
    Возврат: { ISIN: {"pdf_path": <str|None>} }
    PDF располагаются в pdf_dir и именуются '<ISIN>.pdf'
    """
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["TS"]
    ref = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        # row: (N, ISIN, LINK, ...)
        _, isin, _ = (row + (None,))[:3]
        isin = _norm_isin(isin)
        if not isin:
            continue
        pdf_path = os.path.join(pdf_dir, f"{isin}.pdf")
        ref[isin] = {"pdf_path": pdf_path if os.path.isfile(pdf_path) else None}
    return ref


def match_isins(
    isins: List[str],
    ref_stocks: dict,
    ref_bonds: dict,
    ref_struct: dict,
) -> Tuple[List[dict], List[dict], List[dict], List[str]]:
    """
    Сопоставляет ISIN из входного JSON со справочниками.
    Приоритет строго такой (один ISIN → максимум в одну группу):
      1) Акции/ETF
      2) Облигации
      3) Структурные продукты
      4) Иначе — в 'misses'
    Возвращает кортеж списков:
      hits_stocks: list[{"isin","ticker","type"}]
      hits_bonds:  list[{"isin","name"}]
      hits_sp:     list[{"isin","type","pdf_path"}]  # type всегда "СТРУКТУРНЫЙ ПРОДУКТ"
      misses:      list[isin]
    """
    seen = set()
    hits_stocks: List[dict] = []
    hits_bonds: List[dict] = []
    hits_sp: List[dict] = []
    misses: List[str] = []

    for raw in isins:
        isin = (raw or "").strip().upper()
        if not isin or isin in seen:
            continue
        seen.add(isin)

        # 1) Stocks/ETF
        s = ref_stocks.get(isin)
        if s:
            hits_stocks.append({
                "isin": isin,
                "ticker": s.get("ticker", ""),
                "name": s.get("name", ""),
                "type": s.get("type", ""),
            })
            continue

        # 2) Bonds
        b = ref_bonds.get(isin)
        if b:
            hits_bonds.append({
                "isin": isin,
                "name": b.get("name", ""),
            })
            continue

        # 3) Structured products
        sp = ref_struct.get(isin)
        if sp:
            hits_sp.append({
                "isin": isin,
                "type": "СТРУКТУРНЫЙ ПРОДУКТ",
                "pdf_path": sp.get("pdf_path"),
            })
            continue

        # 4) Не найден ни в одном справочнике
        misses.append(isin)

    return hits_stocks, hits_bonds, hits_sp, misses

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

        # Загрузка справочников (только чтение, без сопоставления)
        console.print(f"[green]🔄 Загрузка справочников…[/green]")
        console.print(f"[green]↳ Stocks/ETF:[/green] [bright_cyan]{REF_STOCKS_XLSX}[/bright_cyan]")
        stocks = load_reference_stocks(REF_STOCKS_XLSX)
        console.print(f"[green]   Загружено записей:[/green] [bright_cyan]{len(stocks)}[/bright_cyan]")

        console.print(f"[green]↳ Bonds:[/green] [bright_cyan]{REF_BONDS_XLSX}[/bright_cyan]")
        bonds = load_reference_bonds(REF_BONDS_XLSX)
        console.print(f"[green]   Загружено записей:[/green] [bright_cyan]{len(bonds)}[/bright_cyan]")

        console.print(f"[green]↳ Structured (TS):[/green] [bright_cyan]{REF_SP_XLSX}[/bright_cyan]")
        structured = load_reference_structured(REF_SP_XLSX, REF_SP_PDF_DIR)
        console.print(f"[green]   Загружено записей:[/green] [bright_cyan]{len(structured)}[/bright_cyan]")

        # Сопоставление ISIN по справочникам (без записи на диск)
        hits_stocks, hits_bonds, hits_sp, misses = match_isins(isins, stocks, bonds, structured)

        console.print("[green]🧩 Результат сопоставления:[/green]")
        console.print(f"  Акции/ETF: [bright_cyan]{len(hits_stocks)}[/bright_cyan]")
        console.print(f"  Облигации: [bright_cyan]{len(hits_bonds)}[/bright_cyan]")
        console.print(f"  Структурные продукты: [bright_cyan]{len(hits_sp)}[/bright_cyan]")
        console.print(f"  Неизвестные (noname): [bright_cyan]{len(misses)}[/bright_cyan]")

        # Постоянный предпросмотр результатов (по 3 категориям). Записи не пишутся на диск.
        preview_limit = 20

        def _render_table(title: str, columns: list[str], rows: list[list[str]]):
            if not rows:
                console.print(f"[yellow]{title}: нет записей[/yellow]")
                return
            table = Table(title=title, show_lines=False)
            for col in columns:
                # номер колонки и короткие поля делаем no_wrap для аккуратного вида
                if col in ("№", "ISIN", "Ticker", "Type"):
                    table.add_column(col, no_wrap=True)
                else:
                    table.add_column(col)
            for i, r in enumerate(rows[:preview_limit], 1):
                table.add_row(str(i), *r)
            console.print(table)

        # 1) Предпросмотр: Акции/ETF
        stock_rows = [
            [rec.get("isin", ""), rec.get("ticker", ""), rec.get("name", ""), rec.get("type", "")]
            for rec in hits_stocks
        ]
        _render_table("Предпросмотр stock_etf (будущий JSON)", ["№", "ISIN", "Ticker", "Name", "Type"], stock_rows)

        # 2) Предпросмотр: Облигации
        bond_rows = [
            [rec.get("isin", ""), rec.get("name", "")]
            for rec in hits_bonds
        ]
        _render_table("Предпросмотр bonds (будущий JSON)", ["№", "ISIN", "Name"], bond_rows)

        # 3) Предпросмотр: Структурные продукты
        sp_rows = [
            [rec.get("isin", ""), rec.get("type", "СТРУКТУРНЫЙ ПРОДУКТ")]
            for rec in hits_sp
        ]
        _render_table("Предпросмотр structured (будущий JSON)", ["№", "ISIN", "Type"], sp_rows)

        # Никакой записи файлов и копирования PDF на этом этапе
        console.print("[yellow]Этап 3 завершён: сопоставление выполнено. Запись JSON и копирование PDF будут на следующем этапе.[/yellow]")
        return 0

    except KeyboardInterrupt:
        console.print("\n[red]Операция прервана пользователем[/red]")
        return 1
    except Exception as e:
        console.print(f"[red]❌ Критическая ошибка: {e}[/red]")
        return 1


if __name__ == "__main__":
    sys.exit(main())
