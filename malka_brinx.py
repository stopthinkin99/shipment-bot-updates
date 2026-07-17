"""
malka_brinx.py — Extractor for Malca-Amit and Brinks shipping labels.
All regexes are built against actual OCR output from these label types.
"""

import re
import os
import fitz
import cv2
import numpy as np
import pytesseract
from datetime import datetime

import sys
from pathlib import Path

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
    ("EMBY",         "EMBY"),
    ("UNI CREATION", "UNI"),
    ("UNI-CREATION", "UNI"),
    ("UNI DESIGN",   "UNI"),
    ("FENIX",        "FENIX"),
]


def _route_by_sender(text):
    t = text.upper()
    for keyword, sheet in SENDER_SHEET:
        if keyword in t:
            return sheet
    return ""


def _pdf_to_images(file_path):
    doc = fitz.open(file_path)
    images = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=300)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
        elif pix.n == 3:
            img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        images.append(img)
    doc.close()
    return images


def _ocr_both(img_cv):
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY) if len(img_cv.shape) == 3 else img_cv
    _, binarized = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    psm6  = pytesseract.image_to_string(binarized, config='--oem 3 --psm 6')
    psm11 = pytesseract.image_to_string(binarized, config='--oem 3 --psm 11')
    return psm6 + "\n" + psm11


def is_malca_amit_label(text):
    return bool(re.search(r'MALCA.?AMIT', text, re.IGNORECASE))


def is_brinks_label(text):
    return bool(re.search(r'BRINK', text, re.IGNORECASE))


def is_malca_or_brinks_label(text):
    return is_malca_amit_label(text) or is_brinks_label(text)


def _parse_date(text):
    # Full: Jul 13, 2026
    m = re.search(
        r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.]+(\d{1,2}),?\s+(20\d{2})',
        text, re.IGNORECASE
    )
    if m:
        try:
            return datetime.strptime(
                f"{m.group(1)[:3]} {m.group(2)} {m.group(3)}", '%b %d %Y'
            ).strftime('%Y-%m-%d')
        except Exception:
            pass

    # DD-Mon-YYYY: 13-Jul-2026
    m = re.search(
        r'(\d{1,2})[-./](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[-./](20\d{2})',
        text, re.IGNORECASE
    )
    if m:
        try:
            return datetime.strptime(
                f"{m.group(1)}-{m.group(2)[:3]}-{m.group(3)}", '%d-%b-%Y'
            ).strftime('%Y-%m-%d')
        except Exception:
            pass

    return ""


def _clean_recipient(candidate):
    """Strip phone numbers and junk from a recipient line."""
    # Remove "Phone: xxxxxxxxxx" and everything after
    candidate = re.sub(r'\s*Phone\s*:.*$', '', candidate, flags=re.IGNORECASE)
    # Remove standalone phone numbers at end of line
    candidate = re.sub(r'\s+\d{7,}$', '', candidate)
    return candidate.strip()


def _extract_malca_amit(text, filename, page):
    data = {
        "Source File":     filename,
        "Page":            page,
        "carrier":         "MALCA-AMIT",
        "sheet":           "",
        "date":            "",
        "ship_to":         "",
        "invoice":         "",
        "tracking_number": "",
        "remark":          "",
        "Full Extracted Text": text,
        "Application Run Date and time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    lines = text.split('\n')

    # Sheet: sender name in header block before "To:"
    to_idx = next(
        (i for i, l in enumerate(lines) if re.match(r'^\s*To\s*:', l, re.IGNORECASE)),
        len(lines)
    )
    header_text = '\n'.join(lines[:to_idx])
    data["sheet"] = _route_by_sender(header_text)

    # Date
    data["date"] = _parse_date(text)

    # INV number: "INV: 201053956"
    m = re.search(r'\bINV\s*[:#]?\s*([0-9]{4,})', text, re.IGNORECASE)
    if m:
        data["invoice"] = m.group(1).strip()

    # PO fallback
    if not data["invoice"]:
        m = re.search(r'\bP[\s.]*O\s*[:#]?\s*([A-Z0-9]{4,})', text, re.IGNORECASE)
        if m:
            data["invoice"] = m.group(1).strip()

    # REF fallback
    if not data["invoice"]:
        m = re.search(r'\bREF\s*[:#]?\s*([A-Z0-9]{4,})', text, re.IGNORECASE)
        if m:
            data["invoice"] = m.group(1).strip()

    # Tracking: "Shipment#: 7389729"
    m = re.search(r'Shipment\s*#\s*[:\-]?\s*([0-9]{5,})', text, re.IGNORECASE)
    if m:
        data["tracking_number"] = m.group(1).strip()

    # Recipient: line after "To: <city>" — skip city, take company, strip phone
    for i, line in enumerate(lines):
        if re.match(r'^\s*To\s*:', line, re.IGNORECASE):
            for j in range(i + 1, len(lines)):
                candidate = lines[j].strip()
                if not candidate:
                    continue
                # skip phone-only lines
                if re.match(r'^[\d\s\-\(\)]+$', candidate):
                    continue
                candidate = _clean_recipient(candidate)
                if len(candidate) < 4:
                    continue
                data["ship_to"] = candidate
                break
            break

    return data


def _extract_brinks(text, filename, page):
    data = {
        "Source File":     filename,
        "Page":            page,
        "carrier":         "BRINKS",
        "sheet":           "",
        "date":            "",
        "ship_to":         "",
        "invoice":         "",
        "tracking_number": "",
        "remark":          "",
        "Full Extracted Text": text,
        "Application Run Date and time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    lines = text.split('\n')

    # Sheet: sender appears before "DELIVER TO"
    deliver_idx = next(
        (i for i, l in enumerate(lines) if re.search(r'DELIVER\s*TO', l, re.IGNORECASE)),
        len(lines)
    )
    header_text = '\n'.join(lines[:deliver_idx])
    data["sheet"] = _route_by_sender(header_text)

    # Date
    data["date"] = _parse_date(text)

    # PO supports one full value followed by abbreviated suffixes:
    # PO # 86100087, 89, 90, 107
    m = re.search(
        r"\bP[\s.]*[O0]\s*#?\s*[:\-]?\s*"
        r"([0-9]{5,}(?:\s*[,;/]\s*[0-9]{1,8})+|[A-Z0-9][A-Z0-9,./\-]+)",
        text,
        re.IGNORECASE,
    )
    if m:
        raw_po = m.group(1).strip().rstrip(".,")
        numeric_parts = re.findall(r"\d+", raw_po)
        if numeric_parts and numeric_parts[0].isdigit():
            first = numeric_parts[0]
            expanded = [first]
            for suffix in numeric_parts[1:]:
                if len(suffix) < len(first):
                    expanded.append(first[:-len(suffix)] + suffix)
                else:
                    expanded.append(suffix)
            data["invoice"] = ", ".join(expanded)
        else:
            data["invoice"] = raw_po

    # Tracking: "Tracking No. 11017170336"
    m = re.search(r'Tracking\s*No\.?\s*[:\-]?\s*([0-9]{8,})', text, re.IGNORECASE)
    if m:
        data["tracking_number"] = m.group(1).strip()

    # HAWB fallback
    if not data["tracking_number"]:
        m = re.search(r'HAWB\s*Number\s*[:\-]?\s*([\d\s\-]{6,})', text, re.IGNORECASE)
        if m:
            candidate = re.sub(r'[^0-9]', '', m.group(1))
            if len(candidate) >= 6:
                data["tracking_number"] = candidate

    # Recipient: line after "DELIVER TO:"
    for i, line in enumerate(lines):
        if re.search(r'DELIVER\s*TO\s*:?', line, re.IGNORECASE):
            for j in range(i + 1, len(lines)):
                candidate = lines[j].strip()
                if not candidate:
                    continue
                if re.match(r'^[\d\s\-\(\)]+$', candidate):
                    continue
                if len(candidate) < 4:
                    continue
                candidate = _clean_recipient(candidate)
                if len(candidate) < 4:
                    continue
                data["ship_to"] = candidate
                break
            break

    return data


def extract_malca_brinks_from_file(file_path):
    filename = os.path.basename(file_path)
    images = _pdf_to_images(file_path)
    records = []

    for i, img in enumerate(images):
        text = _ocr_both(img)
        print(f"[MALKA/BRINX DEBUG page {i+1}]\n{text[:500]}\n")

        if is_malca_amit_label(text):
            records.append(_extract_malca_amit(text, filename, i + 1))
        elif is_brinks_label(text):
            records.append(_extract_brinks(text, filename, i + 1))

    return records