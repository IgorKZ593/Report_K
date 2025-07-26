import os
from pathlib import Path

def clear_folder(folder: Path):
    for item in folder.iterdir():
        if item.is_file():
            print(f"Удаляю файл: {item.name}")
            item.unlink()

if __name__ == "__main__":
    folder_path = Path("F:/Python Projets/Report/Data_Backup")
    if folder_path.exists():
        clear_folder(folder_path)
        print("✅ Папка Data_Backup очищена.")
    else:
        print("❌ Папка Data_Backup не найдена.")
