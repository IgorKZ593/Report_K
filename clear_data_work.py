import os
from pathlib import Path

def clear_folder(folder: Path, keep_files: list[str] = [".gitkeep"]):
    for item in folder.iterdir():
        if item.name not in keep_files and item.is_file():
            print(f"Удаляю файл: {item.name}")
            item.unlink()

if __name__ == "__main__":
    folder_path = Path("F:/Python Projets/Report/Data_work")
    if folder_path.exists():
        clear_folder(folder_path)
        print("✅ Папка Data_work очищена.")
    else:
        print("❌ Папка Data_work не найдена.")