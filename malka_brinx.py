"""
malca_amit.py — Extractor for Malca-Amit and Brinks shipping labels.
Detects which carrier by keyword, then extracts fields accordingly.
Called by parser.py when neither standard FedEx nor Zales/UPS pattern matches.
"""

import re
import os
import fitz
import cv2
import numpy as np
import pytesseract
from datetime import datetime

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tessdata"

# ── Sheet routing by sender company keyword ───────────────────────────────
SENDER_SHEET = [
    ("EMBY",        "EMBY"),
    ("UNI CREATION","UNI"),
    ("UNI-CREATION","UNI"),
    ("FENIX",       "FENIX"),
]

# ── Invoice prefix routing (same as parser.py) ────────────────────────────
PREFIX_SHEET = [
    ("2030", "EMBY"),
    ("82",   "FENIX"),
    ("47",   "FENIX"),
    ("10",   "UNI"),
]


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


def _ocr(img_cv):
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY) if len(img_cv.shape) == 3 else img_cv
    _, binarized = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return pytesseract.image_to_string(binarized)


def is_malca_amit_label(text):
    return bool(re.search(r'MALCA.?AMIT', text, re.IGNORECASE))


def is_brinks_label(text):
    return bool(re.search(r'BRINK', text, re.IGNORECASE))


def _route_by_sender(text):
    """Try to determine sheet from sender company name in label header."""
    for keyword, sheet in SENDER_SHEET:
        if keyword.upper() in text.upper():
            return sheet
    return ""


def _route_by_invoice(invoice):
    """Fallback: route by invoice number prefix."""
    for prefix, sheet in PREFIX_SHEET:
        if str(invoice).startswith(prefix):
            return sheet
    return ""


def _parse_date(text):
    """Try multiple date formats found on these labels."""
    # Format: Jul 13, 2026
    m = re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{1,2},?\s+\d{4}', text, re.IGNORECASE)
    if m:
        try:
            return datetime.strptime(m.group(0).replace(',', ''), '%b %d %Y').strftime('%Y-%m-%d')
        except Exception:
            return m.group(0)
    # Format: 13-Jul-2026
    m = re.search(r'(\d{1,2})[.\-](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[.\-](\d{4})', text, re.IGNORECASE)
    if m:
        try:
            return datetime.strptime(m.group(0), '%d-%b-%Y').strftime('%Y-%m-%d')
        except Exception:
            return m.group(0)
    return ""


def _extract_malca_amit(text, filename, page):
    """
    Malca-Amit label fields:
    - Sender block top-left → sheet routing
    - To: <city>\n<COMPANY NAME> → recipient
    - INV: <number> → invoice
    - Shipment#: <number> → tracking
    - Date: <date> → ship date
    """
    data = {
        "Source File":   filename,
        "Page":          page,
        "carrier":       "MALCA-AMIT",
        "sheet":         "",
        "date":          "",
        "ship_to":       "",
        "invoice":       "",
        "tracking_number": "",
        "remark":        "",
        "Full Extracted Text": text,
        "Application Run Date and time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # Sheet: from sender block (top of label before "To:")
    lines = text.split('\n')
    to_idx = next((i for i, l in enumerate(lines) if re.match(r'^\s*To\s*:', l, re.IGNORECASE)), len(lines))
    header_text = '\n'.join(lines[:to_idx])
    data["sheet"] = _route_by_sender(header_text)

    # Date
    data["date"] = _parse_date(text)

    # Recipient: line after "To: <city>" — skip the city line, take company name
    for i, line in enumerate(lines):
        if re.match(r'^\s*To\s*:', line, re.IGNORECASE):
            # The "To:" line itself may have the city — next non-empty line is company
            for j in range(i + 1, len(lines)):
                candidate = lines[j].strip()
                if not candidate:
                    continue
                # skip if it looks like a phone or address number only
                if re.match(r'^[\d\s\-\(\)]+$', candidate):
                    continue
                data["ship_to"] = candidate
                break
            break

    # INV number
    m = re.search(r'\bINV\s*[:#]?\s*([A-Z0-9]{4,})', text, re.IGNORECASE)
    if m:
        data["invoice"] = m.group(1).strip()

    # PO fallback if no INV
    if not data["invoice"]:
        m = re.search(r'\bP[\s.]*O\s*[:#]?\s*([A-Z0-9]{4,})', text, re.IGNORECASE)
        if m:
            data["invoice"] = m.group(1).strip()

    # REF fallback
    if not data["invoice"]:
        m = re.search(r'\bREF\s*[:#]?\s*([A-Z0-9]{4,})', text, re.IGNORECASE)
        if m:
            data["invoice"] = m.group(1).strip()

    # Shipment# → tracking
    m = re.search(r'Shipment\s*#\s*[:\-]?\s*([0-9]{5,})', text, re.IGNORECASE)
    if m:
        data["tracking_number"] = m.group(1).strip()

    # Sheet fallback via invoice prefix
    if not data["sheet"] and data["invoice"]:
        data["sheet"] = _route_by_invoice(data["invoice"])

    return data


def _extract_brinks(text, filename, page):
    """
    Brinks label fields:
    - PO # → invoice
    - HAWB Number / Tracking No. → tracking
    - DELIVER TO: block → recipient
    - Estimated Pu Date → ship date
    - Sender block → sheet routing
    """
    data = {
        "Source File":   filename,
        "Page":          page,
        "carrier":       "BRINKS",
        "sheet":         "",
        "date":          "",
        "ship_to":       "",
        "invoice":       "",
        "tracking_number": "",
        "remark":        "",
        "Full Extracted Text": text,
        "Application Run Date and time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # Sheet from sender (appears before DELIVER TO)
    lines = text.split('\n')
    deliver_idx = next((i for i, l in enumerate(lines) if re.search(r'DELIVER\s*TO', l, re.IGNORECASE)), len(lines))
    header_text = '\n'.join(lines[:deliver_idx])
    data["sheet"] = _route_by_sender(header_text)

    # Date: "Estimated Pu Date: 13-Jul-2026"
    data["date"] = _parse_date(text)

    # Recipient: line after "DELIVER TO:"
    for i, line in enumerate(lines):
        if re.search(r'DELIVER\s*TO\s*:', line, re.IGNORECASE):
            for j in range(i + 1, len(lines)):
                candidate = lines[j].strip()
                if not candidate:
                    continue
                if re.match(r'^[\d\s\-\(\)]+$', candidate):
                    continue
                data["ship_to"] = candidate
                break
            break

    # PO number: "PO #07102026-1,1-2"
    m = re.search(r'PO\s*#\s*([A-Z0-9\-,\.]+)', text, re.IGNORECASE)
    if m:
        data["invoice"] = m.group(1).strip()

    # Tracking: "Tracking No. 11017170336" or "HAWB Number: 1101-7170336"
    m = re.search(r'Tracking\s*No\.?\s*[:\-]?\s*([0-9]{8,})', text, re.IGNORECASE)
    if m:
        data["tracking_number"] = m.group(1).strip()

    if not data["tracking_number"]:
        m = re.search(r'HAWB\s*Number\s*[:\-]?\s*([\d\-]+)', text, re.IGNORECASE)
        if m:
            data["tracking_number"] = re.sub(r'[^0-9]', '', m.group(1))

    # Sheet fallback via invoice prefix
    if not data["sheet"] and data["invoice"]:
        data["sheet"] = _route_by_invoice(data["invoice"])

    return data


def extract_malca_brinks_from_file(file_path):
    """
    Entry point called by parser.py.
    Returns list of dicts (one per page), same shape as other extractors.
    Returns empty list if label is neither Malca-Amit nor Brinks.
    """
    filename = os.path.basename(file_path)
    images = _pdf_to_images(file_path)
    records = []

    for i, img in enumerate(images):
        text = _ocr(img)
        print(f"[MALCA/BRINKS DEBUG page {i+1}]\n{text}\n")  # remove once confirmed

        if is_malca_amit_label(text):
            records.append(_extract_malca_amit(text, filename, i + 1))
        elif is_brinks_label(text):
            records.append(_extract_brinks(text, filename, i + 1))
        # else: not our label type, skip

    return records


def is_malca_or_brinks_label(text):
    """Quick check for parser.py routing decision."""
    return is_malca_amit_label(text) or is_brinks_label(text)