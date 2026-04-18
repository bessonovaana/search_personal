import pathlib

dirs = pathlib.Path('node-1/DATA/ПДнDataset/share')

struct = ['.csv', '.json', '.parquet']
doc = ['.pdf', '.doc', '.docx', '.rtf', '.xls']
web = ['.html']
image = ['.tif', '.jpeg', '.png', '.gif']
video = ['.mp4']

files_by_type = {
    'Структурированные данные': struct,
    'Документы': doc,
    'Веб-страницы': web,
    'Изображения': image,
    'Видео': video
}

for category, extensions in files_by_type.items():
    print(f"\n{category} ({', '.join(extensions)}):")
    found = False
    for f in dirs.rglob('*'):
        if f.is_file() and f.suffix.lower() in extensions:
            print(f"  - {f.name} (размер: {f.stat().st_size} байт)")
            found = True
    if not found:
        print(f"  (нет файлов)")