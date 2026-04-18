import pathlib
from collections import defaultdict
import os
import pathlib
from typing import Dict, List, Optional, Any

# Библиотеки для разных форматов
try:
    from docx import Document  # pip install python-docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False
    print("Warning: python-docx not installed. DOCX files won't be supported.")

try:
    import fitz  # PyMuPDF - pip install PyMuPDF
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False
    print("Warning: PyMuPDF not installed. PDF files won't be supported.")

try:
    from striprtf.striprtf import rtf_to_text  # pip install striprtf
    RTF_AVAILABLE = True
except ImportError:
    RTF_AVAILABLE = False
    print("Warning: striprtf not installed. RTF files won't be supported.")

try:
    import pandas as pd  # pip install pandas openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    print("Warning: pandas not installed. Excel files won't be supported.")

try:
    from pyxtxt import xtxt  # pip install pyxtxt[all]
    PYXtxt_AVAILABLE = True
except ImportError:
    PYXtxt_AVAILABLE = False
    print("Warning: pyxtxt not installed. Install 'pip install pyxtxt[all]' for full support.")


def extract_text_from_document(file_path: pathlib.Path) -> Dict[str, Any]:
    """
    Извлекает текст из документа в зависимости от его расширения.
    
    Поддерживаемые форматы:
    - PDF (.pdf)
    - Word (.docx)
    - Rich Text Format (.rtf)
    - Excel (.xlsx, .xls)
    - Текстовые файлы (.txt)
    
    Args:
        file_path: Path объект с путем к файлу
    
    Returns:
        Словарь с ключами:
        - 'success': bool - успешно ли извлечение
        - 'text': str - извлеченный текст
        - 'error': str - сообщение об ошибке (если есть)
        - 'metadata': dict - метаинформация о файле
    """
    
    result = {
        'success': False,
        'text': '',
        'error': None,
        'metadata': {
            'filename': file_path.name,
            'extension': file_path.suffix.lower(),
            'size_bytes': file_path.stat().st_size if file_path.exists() else 0
        }
    }
    
    if not file_path.exists():
        result['error'] = f"File not found: {file_path}"
        return result
    
    ext = file_path.suffix.lower()
    
    try:
        if ext == '.pdf':
            text = _extract_from_pdf(file_path)
        elif ext == '.docx':
            text = _extract_from_docx(file_path)
        elif ext == '.rtf':
            text = _extract_from_rtf(file_path)
        elif ext in ['.xlsx', '.xls']:
            text = _extract_from_excel(file_path)
        elif ext == '.txt':
            text = _extract_from_txt(file_path)
        else:
            # Пробуем pyxtxt для остальных форматов
            if PYXtxt_AVAILABLE:
                text = xtxt(str(file_path))
            else:
                result['error'] = f"Unsupported format: {ext}"
                return result
        
        result['success'] = True
        result['text'] = text
        result['metadata']['char_count'] = len(text)
        
    except Exception as e:
        result['error'] = str(e)
    
    return result


def _extract_from_pdf(file_path: pathlib.Path) -> str:
    """Извлечение текста из PDF с помощью PyMuPDF (лучший выбор)[citation:3]"""
    if not PDF_AVAILABLE:
        raise ImportError("PyMuPDF (fitz) is required. Install with: pip install PyMuPDF")
    
    text_parts = []
    doc = fitz.open(str(file_path))
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text_parts.append(page.get_text())
    
    doc.close()
    return '\n'.join(text_parts)


def _extract_from_docx(file_path: pathlib.Path) -> str:
    """Извлечение текста из DOCX с сохранением структуры[citation:7]"""
    if not DOCX_AVAILABLE:
        raise ImportError("python-docx is required. Install with: pip install python-docx")
    
    doc = Document(str(file_path))
    text_parts = []
    
    # Извлечение текста из параграфов
    for para in doc.paragraphs:
        if para.text.strip():
            text_parts.append(para.text)
    
    # Извлечение текста из таблиц
    for table in doc.tables:
        for row in table.rows:
            row_text = ' | '.join([cell.text for cell in row.cells if cell.text.strip()])
            if row_text:
                text_parts.append(f"[TABLE] {row_text}")
    
    return '\n'.join(text_parts)


def _extract_from_rtf(file_path: pathlib.Path) -> str:
    """Извлечение текста из RTF[citation:5]"""
    if not RTF_AVAILABLE:
        raise ImportError("striprtf is required. Install with: pip install striprtf")
    
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        rtf_content = f.read()
    
    return rtf_to_text(rtf_content)


def _extract_from_excel(file_path: pathlib.Path) -> str:
    """Извлечение текста из Excel с помощью pandas[citation:4]"""
    if not EXCEL_AVAILABLE:
        raise ImportError("pandas is required. Install with: pip install pandas openpyxl")
    
    # Чтение всех листов Excel
    excel_file = pd.ExcelFile(str(file_path))
    text_parts = []
    
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        
        # Пропускаем пустые DataFrame
        if df.empty:
            continue
        
        # Добавляем заголовок листа
        text_parts.append(f"\n[Sheet: {sheet_name}]")
        
        # Конвертируем DataFrame в текст
        # Заменяем NaN на пустые строки
        df = df.fillna('')
        
        # Добавляем заголовки колонок
        headers = ' | '.join([str(col) for col in df.columns if str(col).strip()])
        if headers:
            text_parts.append(f"Columns: {headers}")
        
        # Добавляем строки данных
        for idx, row in df.iterrows():
            row_values = [str(val).strip() for val in row if str(val).strip()]
            if row_values:
                text_parts.append(' | '.join(row_values))
    
    return '\n'.join(text_parts)


def _extract_from_txt(file_path: pathlib.Path) -> str:
    """Извлечение текста из обычного текстового файла"""
    encodings = ['utf-8', 'cp1251', 'latin-1', 'koi8-r']
    
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    
    # Если ничего не сработало, читаем с ignore ошибок
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()

def process_document_category(files_list: List[Dict]) -> Dict[str, Any]:
    """
    Обрабатывает все документы в категории и извлекает из них текст.
    
    Args:
        files_list: Список словарей с информацией о файлах (из categorized_files['Документы'])
    
    Returns:
        Словарь с результатами обработки
    """
    
    results = {
        'processed': [],
        'failed': [],
        'statistics': {
            'total': len(files_list),
            'successful': 0,
            'failed_count': 0,
            'total_chars': 0,
            'total_size_mb': 0
        }
    }
    
    for file_info in files_list:
        file_path = file_info['path']
        
        print(f"Обработка: {file_info['name']}...")
        
        extraction_result = extract_text_from_document(file_path)
        
        if extraction_result['success']:
            results['processed'].append({
                'filename': file_info['name'],
                'text': extraction_result['text'],
                'char_count': extraction_result['metadata']['char_count'],
                'size_mb': file_info['size'] / (1024 * 1024),
                'extension': file_info['extension']
            })
            results['statistics']['successful'] += 1
            results['statistics']['total_chars'] += extraction_result['metadata']['char_count']
        else:
            results['failed'].append({
                'filename': file_info['name'],
                'error': extraction_result['error'],
                'extension': file_info['extension']
            })
            results['statistics']['failed_count'] += 1
        
        results['statistics']['total_size_mb'] += file_info['size'] / (1024 * 1024)
    
    return results


def save_extracted_texts(results: Dict[str, Any], output_dir: str = 'extracted_texts'):
    """
    Сохраняет извлеченные тексты в отдельные файлы.
    """
    import os
    os.makedirs(output_dir, exist_ok=True)
    
    for item in results['processed']:
        # Создаем безопасное имя файла
        safe_name = item['filename'].replace('/', '_').replace('\\', '_')
        output_path = os.path.join(output_dir, f"{safe_name}.txt")
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(f"Source: {item['filename']}\n")
            f.write(f"Characters: {item['char_count']}\n")
            f.write(f"Size: {item['size_mb']:.2f} MB\n")
            f.write("="*60 + "\n\n")
            f.write(item['text'])
        
        print(f"Сохранен: {output_path}")
    
    # Сохраняем отчет об ошибках
    if results['failed']:
        error_path = os.path.join(output_dir, 'failed_files.txt')
        with open(error_path, 'w', encoding='utf-8') as f:
            for item in results['failed']:
                f.write(f"{item['filename']}: {item['error']}\n")