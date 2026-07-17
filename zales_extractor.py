"""
zales_extractor.py
------------------
Extractor for Zales UPS labels, including UPS "View/Print Label" PDFs where
the actual shipping label is rotated 90 degrees inside a portrait page.
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


def _pdf_to_images(file_path):
    doc = fitz.open(file_path)
    images = []
    try:
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=300)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            if pix.n == 4:
                img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
            elif pix.n == 3:
                img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
            images.append(img)
    finally:
        doc.close()
    return images


def _rotate(img, angle):
    if angle == 90:
        return cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
    if angle == 270:
        return cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    return img


def _ocr_variant(img, psm):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) if len(img.shape) == 3 else img
    gray = cv2.resize(gray, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return pytesseract.image_to_string(binary, config=f"--oem 3 --psm {psm}")


def _ocr_all_rotations(img):
    texts = []
    for angle in (0, 90, 270):
        rotated = _rotate(img, angle)
        for psm in (6, 11):
            texts.append(_ocr_variant(rotated, psm))
    return "\n".join(texts)


def is_zales_label(text):
    return bool(re.search(r"\bZALES\b", str(text or ""), re.IGNORECASE))


def _normalize_ups_tracking(value):
    candidate = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())
    if candidate.startswith(("IZ", "I2", "12", "LZ")):
        candidate = "1Z" + candidate[2:]
    match = re.search(r"1Z[A-Z0-9]{16}", candidate)
    return match.group(0) if match else ""


def _extract_ups_tracking(text):
    match = re.search(
        r"TRACKING\s*#?\s*:?\s*([1IL][Z2][A-Z0-9\s\-]{12,40})",
        text,
        re.IGNORECASE,
    )
    if match:
        tracking = _normalize_ups_tracking(match.group(1))
        if tracking:
            return tracking

    compact = re.sub(r"[^A-Z0-9]", "", text.upper())
    compact = compact.replace("IZ", "1Z").replace("I2", "1Z").replace("LZ", "1Z")
    match = re.search(r"1Z[A-Z0-9]{16}", compact)
    return match.group(0) if match else ""


def _clean_recipient(value):
    value = re.sub(r"\s+", " ", str(value or "")).strip(" .,:;-")
    corrections = {
        "TMARA LEE": "TAMARA LEE",
    }
    return corrections.get(value.upper(), value)


def _extract_recipient(text):
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    for index, line in enumerate(lines):
        if re.search(r"\bSHIP\s*TO\s*:?", line, re.IGNORECASE):
            candidates = []
            for candidate in lines[index + 1:index + 10]:
                upper = candidate.upper()

                if re.search(r"\b(?:UPS|TRACKING|BILLING|SIGNATURE|REFERENCE)\b", upper):
                    break
                if re.fullmatch(r"[\d\s()+\-]+", candidate):
                    continue
                if re.search(r"\b[A-Z]{2}\s+\d{5}(?:-\d{4})?\b", upper):
                    continue
                if re.search(
                    r"\b(?:ST|STREET|AVE|AVENUE|RD|ROAD|BLVD|DR|DRIVE|COVE|LANE|LN)\b",
                    upper,
                ):
                    continue

                letters = len(re.findall(r"[A-Z]", upper))
                digits = len(re.findall(r"\d", upper))
                if letters < 3 or digits > letters:
                    continue

                candidates.append(_clean_recipient(candidate))

            if candidates:
                return candidates[0]

    # Fallback for OCR that drops the literal "SHIP TO:" line.
    for index, line in enumerate(lines):
        if re.fullmatch(r"\d{9,12}", re.sub(r"\D", "", line)):
            for candidate in lines[index + 1:index + 4]:
                upper = candidate.upper()
                if (
                    len(re.findall(r"[A-Z]", upper)) >= 4
                    and not re.search(r"\d", upper)
                    and not re.search(r"\bZALES\b", upper)
                ):
                    return _clean_recipient(candidate)

    return ""


def parse_zales_label(text, filename, page):
    result = {
        "Source File": filename,
        "Page": page,
        "Sheet": "FENIX",
        "Tracking Number": _extract_ups_tracking(text),
        "PO Number": "",
        "Recipient Company": _extract_recipient(text),
        "INV Number": "",
        "Reference": "",
        "CAD": "",
        "Weight": "",
        "Ship Date": datetime.now().strftime("%Y-%m-%d"),
        "Full Extracted Text": text,
        "Application Run Date and time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    match = re.search(
        r"Reference\s*#?\s*1\s*[:\-]?\s*([0-9]{4,})",
        text,
        re.IGNORECASE,
    )
    if match:
        result["PO Number"] = match.group(1)

    match = re.search(
        r"Reference\s*#?\s*2\s*[:\-]?\s*([0-9]{4,})",
        text,
        re.IGNORECASE,
    )
    if match:
        result["Reference"] = match.group(1)

    return result


def extract_zales_from_file(file_path):
    filename = os.path.basename(file_path)
    records = []

    for page_index, image in enumerate(_pdf_to_images(file_path), start=1):
        text = _ocr_all_rotations(image)
        print(f"[ZALES DEBUG page {page_index}]\n{text[:1500]}\n")

        if is_zales_label(text):
            records.append(parse_zales_label(text, filename, page_index))

    return records