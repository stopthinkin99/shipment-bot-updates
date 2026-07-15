import re
import os
import pytesseract
from pdf2image import convert_from_path
from PIL import Image




_pdfinfo = shutil.which("pdfinfo")

if _pdfinfo:
    POPPLER_PATH = str(Path(_pdfinfo).parent)
else:
    POPPLER_PATH = None



# Runtime-safe OCR paths
import sys
from pathlib import Path

_OCR_BASE = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
_TESSERACT = _OCR_BASE / "Tesseract-OCR" / "tesseract.exe"
_TESSDATA = _OCR_BASE / "Tesseract-OCR" / "tessdata"

if not _TESSERACT.exists():
    _TESSERACT = Path(r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe")
if not _TESSDATA.exists():
    _TESSDATA = Path(r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tessdata")

pytesseract.pytesseract.tesseract_cmd = str(_TESSERACT)
os.environ["TESSDATA_PREFIX"] = str(_TESSDATA)

def extract_tracking_number(text):
    # UPS format: 1Z followed by spaced alphanumeric groups
    m = re.search(r'\b(1Z\s*[A-Z0-9]{3}\s*[A-Z0-9]{3}\s*[0-9]{2}\s*[0-9]{4}\s*[0-9]{4})\b', text, re.IGNORECASE)
    if m:
        return re.sub(r'\s+', '', m.group(1))

    # FedEx spaced: 4digits space 4digits space 4digits
    spaced = re.findall(r'\b(\d{4})\s+(\d{4})\s+(\d{4})\b', text)
    if spaced:
        return ''.join(spaced[0])

    # Long continuous numeric run (original)
    compact_text = re.sub(r'(?<=\d)\s+(?=\d)', '', text)
    matches = re.findall(r'\b[0-9]{12,15}\b', compact_text)
    for m in matches:
        if m.startswith(('7', '6', '9')):
            return m
    return matches[0] if matches else ""


def _ocr_image(image_path):
    img = Image.open(image_path)
    return pytesseract.image_to_string(img)


def extract_fields(text):
    data = {
        "Tracking Number": "",
        "PO Number": "",
        "INV Number": ""
    }

    data["Tracking Number"] = extract_tracking_number(text)

    # PO Number
    m = re.search(r'\bP[\s.]*[O0][\s:]*([0-9]{5,})', text, re.IGNORECASE)
    if m:
        data["PO Number"] = m.group(1).strip()
    if not data["PO Number"]:
        for line in text.split('\n'):
            if re.search(r'\bP[\s.]*[O0]\b', line, re.IGNORECASE):
                nums = re.findall(r'\d{5,}', line)
                if nums:
                    data["PO Number"] = nums[0]
                    break

    # INV Number — used when PO is blank
    m = re.search(r'\bINV[\s:]*([0-9]{4,})', text, re.IGNORECASE)
    if m:
        data["INV Number"] = m.group(1).strip()
    if not data["INV Number"]:
        for line in text.split('\n'):
            if re.search(r'\bINV\b', line, re.IGNORECASE):
                nums = re.findall(r'\d{4,}', line)
                if nums:
                    data["INV Number"] = nums[0]
                    break

    return data


def get_po_and_tracking(pdf_path):
    """
    Called by extractor_core.py — returns list of dicts, one per page:
    [{"Tracking Number": "...", "PO Number": "..."}, ...]
    """
    images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
    results = []
    for i, img in enumerate(images):
        img_path = f"{pdf_path}_page{i}.png"
        img.save(img_path)
        text = _ocr_image(img_path)
        results.append(extract_fields(text))
    return results