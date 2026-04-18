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
    total_size = 0
    count = 0
    for f in dirs.rglob('*'):
        if f.is_file() and f.suffix.lower() in extensions:
            count += 1
            total_size += f.stat().st_size
    print(f"{category}: {count} файлов, общий размер {total_size}")