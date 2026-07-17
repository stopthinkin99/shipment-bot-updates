"""
malka_brinx.py
--------------
Extractor for Malca-Amit and Brinks labels.

Brinks improvement:
- detects the barcode area,
- OCRs the narrow strip directly below the barcode,
- supports multiple abbreviated PO values:
      PO # 86100087, 89, 90, 107
  -> 86100087, 86100089, 86100090, 86100107
"""

import os
import re
import sys
from datetime import datetime
from pathlib import Path

import cv2
import fitz
import numpy as np
import pytesseract


_BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
_TESSERACT = _BASE_DIR / "Tesseract-OCR" / "tesseract.exe"
_TESSDATA = _BASE_DIR / "Tesseract-OCR" / "tessdata"

if not _TESSERACT.exists():
    _TESSERACT = Path(r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe")
if not _TESSDATA.exists():
    _TESSDATA = Path(r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tessdata")

pytesseract.pytesseract.tesseract_cmd = str(_TESSERACT)
os.environ["TESSDATA_PREFIX"] = str(_TESSDATA)


SENDER_SHEET = [
    ("EMBY", "EMBY"),
    ("UNI CREATION", "UNI"),
    ("UNI-CREATION", "UNI"),
    ("UNI DESIGN", "UNI"),
    ("FENIX", "FENIX"),
]


def _route_by_sender(text):
    upper = text.upper()
    for keyword, sheet in SENDER_SHEET:
        if keyword in upper:
            return sheet
    return ""


def _pdf_to_images(file_path):
    doc = fitz.open(file_path)
    images = []
    try:
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=350)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            if pix.n == 4:
                img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
            elif pix.n == 3:
                img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
            images.append(img)
    finally:
        doc.close()
    return images


def _ocr(img, psm=6, scale=1.5):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) if len(img.shape) == 3 else img
    gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return pytesseract.image_to_string(binary, config=f"--oem 3 --psm {psm}")


def _ocr_both(img):
    return _ocr(img, 6) + "\n" + _ocr(img, 11)


def is_malca_amit_label(text):
    return bool(re.search(r"MALCA.?AMIT", text, re.IGNORECASE))


def is_brinks_label(text):
    return bool(re.search(r"BRINK", text, re.IGNORECASE))


def is_malca_or_brinks_label(text):
    return is_malca_amit_label(text) or is_brinks_label(text)


def _parse_date(text):
    match = re.search(
        r"(\d{1,2})[-./](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[-./](20\d{2})",
        text,
        re.IGNORECASE,
    )
    if match:
        try:
            return datetime.strptime(
                f"{match.group(1)}-{match.group(2)[:3]}-{match.group(3)}",
                "%d-%b-%Y",
            ).strftime("%Y-%m-%d")
        except ValueError:
            pass

    match = re.search(
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.]+"
        r"(\d{1,2}),?\s+(20\d{2})",
        text,
        re.IGNORECASE,
    )
    if match:
        try:
            return datetime.strptime(
                f"{match.group(1)[:3]} {match.group(2)} {match.group(3)}",
                "%b %d %Y",
            ).strftime("%Y-%m-%d")
        except ValueError:
            pass

    return ""


def _expand_po_list(raw_value):
    numbers = re.findall(r"\d+", raw_value)
    if not numbers:
        return ""

    first = numbers[0]
    expanded = [first]

    for suffix in numbers[1:]:
        if len(suffix) < len(first):
            expanded.append(first[:-len(suffix)] + suffix)
        else:
            expanded.append(suffix)

    return ", ".join(expanded)


def _extract_po_from_text(text):
    patterns = [
        r"\bP[\s.]*[O0]\s*#?\s*[:\-]?\s*"
        r"(\d{5,}(?:\s*[,;/]\s*\d{1,8})+)",
        r"\bP[\s.]*[O0]\s*#?\s*[:\-]?\s*(\d{5,})",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return _expand_po_list(match.group(1))

    return ""


def _po_barcode_strip_text(image):
    height, width = image.shape[:2]
    upper = image[: int(height * 0.55), :]

    gray = cv2.cvtColor(upper, cv2.COLOR_BGR2GRAY)
    inverted = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)[1]

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (35, 5))
    merged = cv2.morphologyEx(inverted, cv2.MORPH_CLOSE, kernel, iterations=2)

    contours, _ = cv2.findContours(
        merged,
        cv2.RETR_EXTERNAL,
        cv2.CHAIN_APPROX_SIMPLE,
    )

    texts = []

    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)

        if w < width * 0.20 or h < 15:
            continue
        if w / max(h, 1) < 3:
            continue

        y1 = max(0, y + h - 10)
        y2 = min(upper.shape[0], y + h + max(100, int(height * 0.08)))
        x1 = max(0, x - 30)
        x2 = min(width, x + w + 30)

        strip = upper[y1:y2, x1:x2]
        if strip.size == 0:
            continue

        for psm in (6, 7, 11, 12):
            texts.append(_ocr(strip, psm=psm, scale=3.0))

    for start_ratio, end_ratio in (
        (0.14, 0.25),
        (0.18, 0.30),
        (0.22, 0.34),
    ):
        strip = image[
            int(height * start_ratio):int(height * end_ratio),
            : int(width * 0.65),
        ]
        if strip.size:
            for psm in (6, 7, 11):
                texts.append(_ocr(strip, psm=psm, scale=3.0))

    return "\n".join(texts)


def _clean_recipient(value):
    value = re.sub(r"\s*Phone\s*:.*$", "", value, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", value).strip(" .,:;-")


def _base_record(filename, page, carrier):
    return {
        "Source File": filename,
        "Page": page,
        "carrier": carrier,
        "sheet": "",
        "date": "",
        "ship_to": "",
        "invoice": "",
        "tracking_number": "",
        "remark": "",
        "Full Extracted Text": "",
        "Application Run Date and time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def _extract_malca_amit(text, filename, page):
    data = _base_record(filename, page, "MALCA-AMIT")
    data["Full Extracted Text"] = text
    data["sheet"] = _route_by_sender(text)
    data["date"] = _parse_date(text)

    for pattern in (
        r"\bINV\s*[:#]?\s*([A-Z0-9\-]{4,})",
        r"\bP[\s.]*O\s*[:#]?\s*([A-Z0-9\-]{4,})",
        r"\bREF\s*[:#]?\s*([A-Z0-9\-]{4,})",
    ):
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data["invoice"] = match.group(1)
            break

    match = re.search(r"Shipment\s*#\s*[:\-]?\s*([0-9]{5,})", text, re.IGNORECASE)
    if match:
        data["tracking_number"] = match.group(1)

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for index, line in enumerate(lines):
        if re.match(r"^To\s*:", line, re.IGNORECASE):
            for candidate in lines[index + 1:index + 8]:
                if not re.fullmatch(r"[\d\s()+\-]+", candidate):
                    data["ship_to"] = _clean_recipient(candidate)
                    break
            break

    return data


def _extract_brinks(text, image, filename, page):
    data = _base_record(filename, page, "BRINKS")
    data["Full Extracted Text"] = text
    data["sheet"] = _route_by_sender(text)
    data["date"] = _parse_date(text)

    data["invoice"] = _extract_po_from_text(text)

    if not data["invoice"]:
        po_text = _po_barcode_strip_text(image)
        print(f"[BRINKS PO OCR]\n{po_text[:1200]}\n")
        data["invoice"] = _extract_po_from_text(po_text)

    match = re.search(
        r"Tracking\s*No\.?\s*[:\-]?\s*([0-9][0-9\s\-]{7,})",
        text,
        re.IGNORECASE,
    )
    if match:
        candidate = re.sub(r"\D", "", match.group(1))
        if len(candidate) >= 8:
            data["tracking_number"] = candidate

    if not data["tracking_number"]:
        match = re.search(
            r"HAWB\s*Number\s*[:\-]?\s*([0-9][0-9\s\-]{6,})",
            text,
            re.IGNORECASE,
        )
        if match:
            candidate = re.sub(r"\D", "", match.group(1))
            if len(candidate) >= 8:
                data["tracking_number"] = candidate

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for index, line in enumerate(lines):
        if re.search(r"DELIVER\s*TO\s*:?", line, re.IGNORECASE):
            for candidate in lines[index + 1:index + 8]:
                if re.fullmatch(r"[\d\s()+\-]+", candidate):
                    continue
                if re.search(r"\b(?:UNIT|CENTRE|CENTER|STREET|ROAD|PHONE|CONTACT)\b", candidate, re.IGNORECASE):
                    continue
                cleaned = _clean_recipient(candidate)
                if len(cleaned) >= 4:
                    data["ship_to"] = cleaned
                    break
            break

    return data


def extract_malca_brinks_from_file(file_path):
    filename = os.path.basename(file_path)
    records = []

    for page_index, image in enumerate(_pdf_to_images(file_path), start=1):
        text = _ocr_both(image)
        print(f"[MALKA/BRINKS DEBUG page {page_index}]\n{text[:1500]}\n")

        if is_malca_amit_label(text):
            records.append(_extract_malca_amit(text, filename, page_index))
        elif is_brinks_label(text):
            records.append(_extract_brinks(text, image, filename, page_index))

    return records