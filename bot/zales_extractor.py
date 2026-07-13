import re
import os
import fitz
import cv2
import numpy as np
import pytesseract
from datetime import datetime

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tessdata"


def _pdf_to_images(file_path):
    doc = fitz.open(file_path)
    images = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=300)
        img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
        if pix.n == 4:
            img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGBA2BGR)
        elif pix.n == 3:
            img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
        else:
            img_cv = img_array
        images.append(img_cv)
    doc.close()
    return images


def _ocr(img_cv):
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY) if len(img_cv.shape) == 3 else img_cv
    _, binarized = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return pytesseract.image_to_string(binarized)


def is_zales_label(text):
    return bool(re.search(r'\bZALES\b', text, re.IGNORECASE))


def _extract_ups_tracking(text):
    # Strategy 1: clean 1Z format fully intact
    m = re.search(r'(1Z\s*[A-Z0-9]{3}\s*[A-Z0-9]{3}\s*[0-9]{2}\s*[0-9]{4}\s*[0-9]{4})', text, re.IGNORECASE)
    if m:
        return re.sub(r'\s+', '', m.group(1))

    # Strategy 2: TRACKING #: line — grab everything after the colon
    m = re.search(r'TRACKING\s*#\s*:\s*([A-Z0-9][\w\s]{6,30})', text, re.IGNORECASE)
    if m:
        return re.sub(r'\s+', '', m.group(1).strip())

    # Strategy 3: OCR mangled "UPS GROUND | 1047 0932" — grab digits near UPS GROUND
    # The label's actual tracking ends in those digit groups even if 1Z is lost
    m = re.search(r'(?:UPS\s*G\s*ROUND|UPS\s*GROUND)[^\n]*?([0-9]{4}\s+[0-9]{4})', text, re.IGNORECASE)
    if m:
        # We only have the tail — prefix with known UPS marker so it's identifiable
        return "UPS-" + re.sub(r'\s+', '', m.group(1))

    return ""


def _extract_recipient(text):
    """
    After SHIP TO: skip:
    - lines that are only digits/phone numbers
    - lines shorter than 4 chars
    - lines that look garbled (too many digits relative to letters)
    Take the first clean alphabetic name line, then check if the line
    after it is the same name repeated (common on UPS labels) — if so,
    use the second occurrence since it tends to be cleaner OCR.
    """
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if re.match(r'^\s*SHIP\s*TO\s*:', line, re.IGNORECASE):
            candidates = []
            for j in range(i + 1, len(lines)):
                candidate = lines[j].strip()
                if not candidate:
                    continue
                # skip pure phone/digit lines
                if re.match(r'^[\d\s\-\(\)]+$', candidate):
                    continue
                # skip very short lines
                if len(candidate) < 4:
                    continue
                # skip lines where digits dominate
                digits = len(re.findall(r'\d', candidate))
                letters = len(re.findall(r'[a-zA-Z]', candidate))
                if digits > letters:
                    continue
                candidates.append(candidate)
                if len(candidates) == 2:
                    break

            if not candidates:
                return ""
            # If both candidates look like the same name, return the second
            # (repeat of name on UPS labels is usually cleaner OCR)
            if (len(candidates) == 2 and
                    candidates[0].upper().replace(' ', '') in candidates[1].upper().replace(' ', '') or
                    candidates[1].upper().replace(' ', '') in candidates[0].upper().replace(' ', '')):
                return candidates[1]
            return candidates[0]

    return ""


def parse_zales_label(text, filename, page):
    data = {
        "Source File": filename,
        "Page": page,
        "Sheet": "FENIX",
        "Tracking Number": "",
        "PO Number": "",
        "Recipient Company": "",
        "INV Number": "",
        "Reference": "",
        "CAD": "",
        "Weight": "",
        "Ship Date": datetime.now().strftime("%Y-%m-%d"),
        "Full Extracted Text": text,
        "Application Run Date and time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    data["Tracking Number"] = _extract_ups_tracking(text)
    data["Recipient Company"] = _extract_recipient(text)

    # Reference #1 → PO field
    m = re.search(r'Reference\s*#\s*1[\s:]*([0-9]+)', text, re.IGNORECASE)
    if m:
        data["PO Number"] = m.group(1).strip()

    # Reference #2 → Reference field
    m = re.search(r'Reference\s*#\s*2[\s:]*([0-9]+)', text, re.IGNORECASE)
    if m:
        data["Reference"] = m.group(1).strip()

    return data


def extract_zales_from_file(file_path):
    filename = os.path.basename(file_path)
    images = _pdf_to_images(file_path)
    records = []
    for i, img in enumerate(images):
        text = _ocr(img)
        print(f"[ZALES DEBUG page {i+1}]\n{text}\n")   # remove once confirmed working
        if is_zales_label(text):
            records.append(parse_zales_label(text, filename, i + 1))
    return records