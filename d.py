#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import csv
import json
import argparse
import logging
import subprocess
import tempfile
import gc
import datetime
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed
from collections import defaultdict
from typing import Dict, List, Optional
import pandas as pd
import pdfplumber
from docx import Document
from bs4 import BeautifulSoup
from PIL import Image
import pytesseract
from striprtf.striprtf import rtf_to_text
import olefile

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfplumber").setLevel(logging.ERROR)

SUPPORTED_EXTENSIONS = {
    ".csv", ".json", ".parquet", ".pdf", ".doc", ".docx", ".rtf",
    ".xls", ".xlsx", ".html", ".htm", ".tif", ".tiff", ".jpeg",
    ".jpg", ".png", ".gif", ".mp4"
}

MAX_FILE_SIZE_MB = 500
MAX_PDF_PAGES = 50
THRESHOLD_LARGE = 100

PD_PATTERNS = {
    "standard_fio": re.compile(r'\b[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}\b'),
    "standard_phone": re.compile(r'(?:\+7|8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}'),
    "standard_email": re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b'),
    "standard_dob": re.compile(r'\b(?:0[1-9]|[12][0-9]|3[01])[./-](?:0[1-9]|1[0-2])[./-](?:19|20)\d{2}\b'),
    "standard_address": re.compile(r'(?:г\.|город|ул\.|улица|пр\.|проспект|д\.|дом|кв\.|квартира|обл\.|область|р-н|район)\s+[А-ЯЁа-яё0-9\s\.\-\/,]{5,}'),
    "state_passport": re.compile(r'\b\d{4}\s?\d{6}\b'),
    "state_snils": re.compile(r'\b\d{3}[\s\-]?\d{3}[\s\-]?\d{3}[\s\-]?\d{2}\b'),
    "state_inn": re.compile(r'\b\d{10}\b|\b\d{12}\b'),
    "state_driver_license": re.compile(r'\b\d{2}\s?[A-ZА-ЯЁ]{2}\s?\d{6}\b', re.IGNORECASE),
    "state_mrz": re.compile(r'P<.+?\n.+?<<', re.MULTILINE),
    "payment_card": re.compile(r'\b\d{16}\b'),
    "payment_account": re.compile(r'\b\d{20}\b'),
    "payment_bik": re.compile(r'\b04\d{7}\b'),
    "payment_cvv": re.compile(r'\b\d{3}\b'),
    "biometric": re.compile(r'(?:отпечаток|палец|радужная оболочка|сетчатка|голосовой образ|биометрия|распознавание лица|скан лица)', re.IGNORECASE),
    "special_health": re.compile(r'(?:диагноз|заболевание|анализ|рентген|мрт|группа крови|инвалидность|справка|история болезни)', re.IGNORECASE),
    "special_religion": re.compile(r'(?:вероисповедание|религиозные убеждения|церковь|мечеть|синагога|конфессия|прихожанин|политические убеждения)', re.IGNORECASE),
    "special_race": re.compile(r'(?:национальность|расовая принадлежность|этническое происхождение|коренной|диаспора)', re.IGNORECASE),
}

CATEGORY_GROUPS = {
    "special": {"special_health", "special_religion", "special_race"},
    "biometric": {"biometric"},
    "payment": {"payment_card", "payment_account", "payment_bik", "payment_cvv"},
    "state": {"state_passport", "state_snils", "state_inn", "state_driver_license", "state_mrz"},
    "standard": {"standard_fio", "standard_phone", "standard_email", "standard_dob", "standard_address"}
}

def luhn_check(card: str) -> bool:
    nums = [int(d) for d in card if d.isdigit()]
    if len(nums) != 16:
        return False
    total = sum(nums[::-1][i] * 2 if i % 2 else nums[::-1][i] for i in range(16))
    total -= sum(9 for x in nums[::-1] if x * 2 > 9)
    return total % 10 == 0

def snils_check(snils: str) -> bool:
    nums = [int(d) for d in snils if d.isdigit()]
    if len(nums) != 11:
        return False
    total = sum(nums[i] * (9 - i) for i in range(9))
    check = nums[9] * 10 + nums[10]
    if total < 100:
        return total == check
    elif total in (100, 101):
        return check == 0
    return (total % 101 == check) or (total % 101 == 100 and check == 0)

def extract_text_fast(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    size_mb = file_path.stat().st_size / (1024**2)
    if size_mb > MAX_FILE_SIZE_MB:
        return ""

    try:
        if ext == ".csv":
            return pd.read_csv(file_path, nrows=10000 if size_mb > 10 else None).to_string(index=False, header=False)
        elif ext == ".json":
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            text = json.dumps(data, ensure_ascii=False, default=str)
            return text[:500000] if size_mb > 10 else text
        elif ext == ".parquet":
            return pd.read_parquet(file_path, columns=None).head(10000).to_string(index=False, header=False)
        elif ext == ".pdf":
            with pdfplumber.open(file_path) as pdf:
                pages = pdf.pages[:MAX_PDF_PAGES]
                return "\n".join(page.extract_text() or "" for page in pages)
        elif ext == ".docx":
            doc = Document(file_path)
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())[:500000]
        elif ext == ".doc":
            try:
                res = subprocess.run(["antiword", "-m UTF-8", str(file_path)], capture_output=True, text=True, timeout=15)
                if res.returncode == 0:
                    return res.stdout[:500000]
            except Exception:
                pass
            try:
                ole = olefile.OleFileIO(file_path)
                text = ole.openstream("WordDocument").read().decode("utf-16-le", errors="ignore")
                return text[:500000]
            except Exception:
                return ""
        elif ext == ".rtf":
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                return rtf_to_text(f.read())[:500000]
        elif ext in (".xls", ".xlsx"):
            return pd.read_excel(file_path).head(10000).to_string(index=False, header=False)
        elif ext in (".html", ".htm"):
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                soup = BeautifulSoup(f.read(), "html.parser")
                return soup.get_text(separator="\n", strip=True)[:500000]
        elif ext in (".tif", ".tiff", ".jpeg", ".jpg", ".png", ".gif"):
            if file_path.stat().st_size < 10240:
                return ""
            img = Image.open(file_path)
            return pytesseract.image_to_string(img, config='--psm 6 -l rus+eng --oem 3')
        elif ext == ".mp4":
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
                try:
                    subprocess.run(["ffmpeg", "-ss", "00:00:01", "-i", str(file_path), "-frames:v", "1", "-q:v", "2", tmp.name],
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30)
                    if os.path.getsize(tmp.name) > 0:
                        return pytesseract.image_to_string(Image.open(tmp.name), config='--psm 6 -l rus+eng --oem 3')
                except Exception:
                    return ""
                finally:
                    if os.path.exists(tmp.name):
                        os.unlink(tmp.name)
            return ""
        else:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()[:500000]
    except Exception:
        return ""

def detect_pd(text: str) -> Dict[str, int]:
    if not text or len(text) < 10:
        return {}
    counts = defaultdict(int)
    for cat, pattern in PD_PATTERNS.items():
        matches = pattern.findall(text)
        if cat == "payment_card":
            matches = [m for m in matches if luhn_check(m)]
        elif cat == "state_snils":
            matches = [m for m in matches if snils_check(m)]
        counts[cat] = len(matches)
    return {k: v for k, v in counts.items() if v > 0}

def classify_uz(categories: Dict[str, int]) -> str:
    has_special = any(c in CATEGORY_GROUPS["special"] for c in categories)
    has_biometric = "biometric" in categories
    has_payment = any(c in CATEGORY_GROUPS["payment"] for c in categories)
    has_state = any(c in CATEGORY_GROUPS["state"] for c in categories)
    has_standard = any(c in CATEGORY_GROUPS["standard"] for c in categories)

    state_total = sum(categories.get(c, 0) for c in CATEGORY_GROUPS["state"])
    standard_total = sum(categories.get(c, 0) for c in CATEGORY_GROUPS["standard"])

    if has_special or has_biometric:
        return "УЗ-1"
    if has_payment or (has_state and state_total > THRESHOLD_LARGE):
        return "УЗ-2"
    if (has_state and state_total <= THRESHOLD_LARGE) or (has_standard and standard_total > THRESHOLD_LARGE):
        return "УЗ-3"
    return "УЗ-4"

def process_file(file_path: Path, root_dir: Path) -> Optional[Dict]:
    if not file_path.is_file() or file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return None
    
    stat = file_path.stat()
    size = stat.st_size
    mtime = datetime.datetime.fromtimestamp(stat.st_mtime)
    time_str = f"{mtime.strftime('%b').lower()} {mtime.day} {mtime.strftime('%H:%M')}"
    
    try:
        name = str(file_path.relative_to(root_dir))
    except ValueError:
        name = file_path.name
    
    text = extract_text_fast(file_path)
    if not text:
        return None
        
    categories = detect_pd(text)
    if not categories:
        return None
        
    uz = classify_uz(categories)
    gc.collect()
    
    return {
        "meta": {"size": size, "time": time_str, "name": name},
        "report": {
            "Путь": name,
            "Категории": ", ".join([f"{k}: {v}" for k, v in categories.items()]),
            "Кол-во": sum(categories.values()),
            "УЗ": uz,
            "Формат": file_path.suffix.lower(),
            "Рекомендации": "AES-256 + RBAC + Аудит" if uz in ["УЗ-1", "УЗ-2"] else "Контроль доступа + Сверка"
        }
    }

def write_reports(all_results: List[Dict], result_path: Path, report_path: Path):
    metas = []
    reports = []
    for res in all_results:
        if res["meta"]["size"] and res["meta"]["time"] and res["meta"]["name"]:
            metas.append(res["meta"])
        if res["report"]["Путь"] and res["report"]["УЗ"]:
            reports.append(res["report"])
        
    with open(result_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["size", "time", "name"])
        for m in sorted(metas, key=lambda x: x["name"]):
            w.writerow([m["size"], m["time"], m["name"]])
            
    with open(report_path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Путь", "Категории", "Кол-во", "УЗ", "Формат", "Рекомендации"])
        for r in sorted(reports, key=lambda x: x["Путь"]):
            w.writerow([r["Путь"], r["Категории"], r["Кол-во"], r["УЗ"], r["Формат"], r["Рекомендации"]])
            
    logging.info(f"Файлы сохранены: {result_path}, {report_path}")

def main():
    parser = argparse.ArgumentParser(description="Сканер ПДн (152-ФЗ)")
    parser.add_argument("input_dir", type=Path, default='ПДнDataset/share', help="Путь к директории")
    parser.add_argument("-w", "--workers", type=int, default=4, help="Число процессов")
    args = parser.parse_args()

    workers = args.workers if args.workers > 0 else min(8, os.cpu_count())
    logging.info(f"Сканирование: {args.input_dir.resolve()} | Workers: {workers}")

    files = [f for f in args.input_dir.rglob("*") if f.suffix.lower() in SUPPORTED_EXTENSIONS and f.is_file()]
    results = []
    processed = 0

    with ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(process_file, f, args.input_dir): f for f in files}
        for future in as_completed(futures):
            try:
                res = future.result(timeout=120)
                if res:
                    results.append(res)
                processed += 1
                if processed % 50 == 0:
                    logging.info(f"Обработано: {processed}/{len(files)}")
            except Exception as e:
                logging.debug(f"Ошибка: {e}")

    write_reports(results, Path("result.csv"), Path("report.csv"))
    logging.info(f"Найдено файлов с ПДн: {len(results)} / {len(files)}")

if __name__ == "__main__":
    main()