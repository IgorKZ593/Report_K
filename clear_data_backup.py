#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import shutil
from pathlib import Path

try:
    from rich.console import Console
    from rich import print
except ImportError:
    import sys
    os.system(f'"{sys.executable}" -m pip install rich')
    from rich.console import Console
    from rich import print

console = Console()

# Папка с резервами
DATA_BACKUP = Path(r"F:\Python Projets\Report\Data_Backup")

def cleanup_backup():
    if not DATA_BACKUP.exists():
        console.print(f"[red]❌ Папка не найдена:[/red] [bright_cyan]{DATA_BACKUP}[/bright_cyan]")
        return

    removed_any = False
    for item in DATA_BACKUP.iterdir():
        try:
            if item.is_file():
                item.unlink()
                console.print(f"[bright_cyan]Удалён файл: {item.name}[/bright_cyan]")
            elif item.is_dir():
                shutil.rmtree(item)
                console.print(f"[bright_cyan]Удалена папка: {item.name}[/bright_cyan]")
            removed_any = True
        except Exception as e:
            console.print(f"[red]⚠ Ошибка удаления {item}: {e}[/red]")

    if not removed_any:
        console.print("[yellow]Папка Data_Backup пуста[/yellow]")
    else:
        console.print("[green]✅ Очистка завершена[/green]")

if __name__ == "__main__":
    cleanup_backup()
