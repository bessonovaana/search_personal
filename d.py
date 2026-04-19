import os
import re
import csv
import json
import argparse
from pathlib import Path
from collections import defaultdict
from typing import Dict, List
import pandas as pd
import pdfplumber
from docx import Document
from bs4 import BeautifulSoup
from PIL import Image
import pytesseract

SUPPORTED_EXTENSIONS = {
    ".csv", ".json", ".parquet", ".pdf", ".doc", ".docx", ".rtf", ".xls", ".xlsx",
    ".html",  ".tif", ".tiff", ".jpeg", ".jpg", ".png", ".gif", ".mp4"
}


THRESHOLD_LARGE = 100


PD_PATTERNS = {
    "standard_fio": re.compile(r'\b[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}\b'),
    "standard_phone": re.compile(r'(?:\+7|8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}'),
    "standard_email": re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'),
    "standard_dob": re.compile(r'\b(?:0[1-9]|[12][0-9]|3[01])[./-](?:0[1-9]|1[0-2])[./-](?:19|20)\d{2}\b'),
    "standard_address": re.compile(r'(?:г\.|город|ул\.|улица|пр\.|проспект|д\.|дом|кв\.|квартира)\s+[А-ЯЁа-яё0-9\s\.\-]+', re.IGNORECASE),
    
    "state_passport": re.compile(r'\b\d{4}\s?\d{6}\b'),
    "state_snils": re.compile(r'\b\d{3}[\s\-]?\d{3}[\s\-]?\d{3}[\s\-]?\d{2}\b'),
    "state_inn": re.compile(r'\b\d{10}\b|\b\d{12}\b'),
    "state_driver_license": re.compile(r'\b\d{2}\s?[A-ZА-ЯЁ]{2}\s?\d{6}\b'),
    
    "payment_card": re.compile(r'\b\d{16}\b'),
    "payment_account": re.compile(r'\b\d{20}\b'),
    "payment_bik": re.compile(r'\b04\d{7}\b'),
    "payment_cvv": re.compile(r'\b\d{3}\b'), 
    
    "biometric": re.compile(r'(?:отпечаток|палец|радужная оболочка|сетчатка|голосовой образ|биометрия|распознавание лица)', re.IGNORECASE),
    "special_health": re.compile(r'(?:диагноз|заболевание|анализ|рентген|мрт|группа крови|инвалидность)', re.IGNORECASE),
    "special_religion": re.compile(r'(?:вероисповедание|религиозные убеждения|церковь|мечеть|синагога)', re.IGNORECASE),
    "special_race": re.compile(r'(?:национальность|расовая принадлежность|этническое происхождение)', re.IGNORECASE),
}


CATEGORY_GROUPS = {
    "special": {"special_health", "special_religion", "special_race"},
    "biometric": {"biometric"},
    "payment": {"payment_card", "payment_account", "payment_bik", "payment_cvv"},
    "state": {"state_passport", "state_snils", "state_inn", "state_driver_license"},
    "standard": {"standard_fio", "standard_phone", "standard_email", "standard_dob", "standard_address"}
}


def luhn_check(card: str) -> bool:
    """Проверка номера банковской карты алгоритмом Луна."""
    nums = [int(d) for d in card if d.isdigit()]
    if len(nums) != 16: return False
    total = sum(nums[::-1][i] * 2 if i % 2 else nums[::-1][i] for i in range(16))
    total -= sum(9 for x in nums[::-1] if x * 2 > 9)
    return total % 10 == 0

def mask_value(val: str) -> str:
   
    if len(val) <= 4: return "***"
    return val[:2] + "*" * (len(val) - 4) + val[-2:]


def extract_text(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    try:
        if ext in {".csv", ".json", ".parquet"}:
            df = pd.read_csv(file_path) if ext == ".csv" else (pd.read_json(file_path) if ext == ".json" else pd.read_parquet(file_path))
            return df.astype(str).to_csv(index=False, header=False)
        elif ext in {".pdf"}:
            with pdfplumber.open(file_path) as pdf:
                return "\n".join(page.extract_text() or "" for page in pdf.pages)
        elif ext in {".docx"}:
            doc = Document(file_path)
            return "\n".join(p.text for p in doc.paragraphs)
        elif ext in {".xls", ".xlsx"}:
            df = pd.read_excel(file_path)
            return df.astype(str).to_csv(index=False, header=False)
        elif ext in {".html"}:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                soup = BeautifulSoup(f.read(), "html.parser")
                return soup.get_text(separator=" ", strip=True)
        elif ext in {".tif", ".tiff", ".jpeg", ".jpg", ".png", ".gif"}:
            img = Image.open(file_path)
            return pytesseract.image_to_string(img, lang="rus+eng")
        elif ext in {".mp4"}:
            return "[АУДИО/ВИДЕО КОНТЕНТ] Требуется отдельная транскрибация"
        else:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
    except Exception as e:
        return f"[ОШИБКА ЧТЕНИЯ] {str(e)}"


def detect_pd(text: str) -> Dict[str, int]:
    """Возвращает количество найденных ПДн по категориям."""
    counts = defaultdict(int)
    for cat, pattern in PD_PATTERNS.items():
        matches = pattern.findall(text)
        if cat == "payment_card":
            matches = [m for m in matches if luhn_check(m)]
        counts[cat] = len(matches)
    return dict(counts)

def classify_uz(categories: Dict[str, int]) -> str:
    has_special = any(c in CATEGORY_GROUPS["special"] for c, v in categories.items() if v > 0)
    has_biometric = any(c in CATEGORY_GROUPS["biometric"] for c, v in categories.items() if v > 0)
    has_payment = any(c in CATEGORY_GROUPS["payment"] for c, v in categories.items() if v > 0)
    has_state = any(c in CATEGORY_GROUPS["state"] for c, v in categories.items() if v > 0)
    has_standard = any(c in CATEGORY_GROUPS["standard"] for c, v in categories.items() if v > 0)

    state_total = sum(v for c, v in categories.items() if c in CATEGORY_GROUPS["state"])
    standard_total = sum(v for c, v in categories.items() if c in CATEGORY_GROUPS["standard"])

    is_state_large = state_total > THRESHOLD_LARGE
    is_standard_large = standard_total > THRESHOLD_LARGE

    if has_special or has_biometric: return "УЗ-1"
    if has_payment or (has_state and is_state_large): return "УЗ-2"
    if (has_state and not is_state_large) or (has_standard and is_standard_large): return "УЗ-3"
    return "УЗ-4"



def generate_report(results: List[Dict], fmt: str, out_path: Path):
    if not results:
        print("Файлы с ПДн не найдены.")
        return

    if fmt == "csv":
        with open(out_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Путь", "Категории ПДн", "Количество", "УЗ", "Формат", "Рекомендации"])
            for r in results:
                cats = ", ".join([f"{k}: {v}" for k, v in r["categories"].items() if v > 0])
                rec = "Шифрование + RBAC" if r["uz"] in ["УЗ-1", "УЗ-2"] else "Контроль доступа + аудит"
                writer.writerow([r["path"], cats, sum(r["categories"].values()), r["uz"], r["ext"], rec])
    print(f"Отчет сохранен: {out_path}")


def main():
    parser = argparse.ArgumentParser(description="Сканер ПДн для корпоративных хранилищ")
    parser.add_argument("input_dir", type=Path, help="Путь к директории для сканирования")
    parser.add_argument("-o", "--output", type=Path, default="pd_report.csv", help="Путь к файлу отчета")
    parser.add_argument("-f", "--format", choices=["csv", "json", "md"], default="csv", help="Формат отчета")
    parser.add_argument("-t", "--threshold", type=int, default=100, help="Порог 'большого объема' для УЗ")
    args = parser.parse_args()

    global THRESHOLD_LARGE
    THRESHOLD_LARGE = args.threshold

    results = []
    print(f"Сканирование директории: {args.input_dir.resolve()}")
    
    for file_path in args.input_dir.rglob("*"):
        if not file_path.is_file() or file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
            
        print(f"Обработка: {file_path.name}")
        text = extract_text(file_path)
        if "[ОШИБКА ЧТЕНИЯ]" in text:
            continue
            
        categories = detect_pd(text)
        if sum(categories.values()) == 0:
            continue
            
        uz = classify_uz(categories)
        results.append({
            "path": str(file_path.relative_to(args.input_dir)),
            "categories": {k: v for k, v in categories.items() if v > 0},
            "uz": uz,
            "ext": file_path.suffix.lower()
        })

    generate_report(results, args.format, args.output)
    print(f"Итого файлов с ПДн: {len(results)}")

if __name__ == "__main__":
    main()