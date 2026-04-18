import pathlib
from collections import defaultdict
import pandas as pd

def struct_data(files):

    if not files:
        print("  ⚠️ В категории нет файлов.")
        return

    print(f"\n{'='*65}")
    print(f"📊 Структурированные данные: {len(files)} файлов")
    print(f"{'='*65}")
    
    for idx, f_info in enumerate(files, start=1):
        print(f"\n🔹 Файл #{idx}")
        print(f"   Имя:        {f_info['name']}")
        print(f"   Полный путь:{f_info['path']}")
        print(f"   Расширение: {f_info['extension']}")
        print(f"   Род. папка: {f_info['parent']}")
        print("-" * 55)
    

dirs = pathlib.Path('ПДнDataset/share')

files_by_type = {
    'Структурированные данные': ['.csv', '.json', '.parquet'],
    'Документы': ['.pdf', '.doc', '.docx', '.rtf', '.xls'],
    'Веб-страницы': ['.html'],
    'Изображения': ['.tif', '.jpeg', '.png', '.gif'],
    'Видео': ['.mp4']
}

categorized_files = {
    'Структурированные данные': [],
    'Документы': [],
    'Веб-страницы': [],
    'Изображения': [],
    'Видео': [],
}

for f in dirs.rglob('*'):
    if f.is_file():
        file_ext = f.suffix.lower()
        file_info = {
            'path': f,
            'name': f.name,
            'size': f.stat().st_size,
            'extension': file_ext,
            'parent': f.parent
        }
        
        categorized = False
        for category, extensions in files_by_type.items():
            if file_ext in extensions:
                categorized_files[category].append(file_info)
                categorized = True
                break

for category, files in categorized_files.items():
    print(f"{category}: {len(files)} файлов")
    if category=='Структурированные данные':
        struct_data(files)
        

