import os
import re
import sys
from pathlib import Path

import pytesseract
from pdf2image import convert_from_path
from PIL import Image


# ------------------------------------------------------------------ #
#  RUNTIME PATHS
# ------------------------------------------------------------------ #
_BASE_DIR = (
    Path(sys.executable).parent
    if getattr(sys, "frozen", False)
    else Path(__file__).parent
)

# Bundled Poppler first
_BUNDLED_POPPLER = _BASE_DIR / "poppler" / "bin"

# Development fallback
_DEV_POPPLER = Path(
    r"C:\Users\aayan.boradia\Downloads\poppler-26.02.0\Library\bin"
)

if (_BUNDLED_POPPLER / "pdfinfo.exe").exists():
    POPPLER_PATH = str(_BUNDLED_POPPLER)
elif (_DEV_POPPLER / "pdfinfo.exe").exists():
    POPPLER_PATH = str(_DEV_POPPLER)
else:
    POPPLER_PATH = None


# Bundled Tesseract first
_TESSERACT = _BASE_DIR / "Tesseract-OCR" / "tesseract.exe"
_TESSDATA = _BASE_DIR / "Tesseract-OCR" / "tessdata"

# Development fallbacks
if not _TESSERACT.exists():
    _TESSERACT = Path(
        r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe"
    )

if not _TESSDATA.exists():
    _TESSDATA = Path(
        r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tessdata"
    )

pytesseract.pytesseract.tesseract_cmd = str(_TESSERACT)
os.environ["TESSDATA_PREFIX"] = str(_TESSDATA)


def extract_tracking_number(text):
    # UPS format: 1Z followed by spaced alphanumeric groups
    match = re.search(
        r"\b(1Z\s*[A-Z0-9]{3}\s*[A-Z0-9]{3}\s*"
        r"[0-9]{2}\s*[0-9]{4}\s*[0-9]{4})\b",
        text,
        re.IGNORECASE,
    )
    if match:
        return re.sub(r"\s+", "", match.group(1))

    # FedEx spaced: 4 digits + 4 digits + 4 digits
    spaced = re.findall(r"\b(\d{4})\s+(\d{4})\s+(\d{4})\b", text)
    if spaced:
        return "".join(spaced[0])

    compact_text = re.sub(r"(?<=\d)\s+(?=\d)", "", text)
    matches = re.findall(r"\b[0-9]{12,15}\b", compact_text)

    for number in matches:
        if number.startswith(("7", "6", "9")):
            return number

    return matches[0] if matches else ""


def _ocr_image(image_path):
    image = Image.open(image_path)
    try:
        return pytesseract.image_to_string(image)
    finally:
        image.close()


def extract_fields(text):
    data = {
        "Tracking Number": "",
        "PO Number": "",
        "INV Number": "",
    }

    data["Tracking Number"] = extract_tracking_number(text)

    # PO Number
    match = re.search(
        r"\bP[\s.]*[O0][\s:]*([0-9]{5,})",
        text,
        re.IGNORECASE,
    )
    if match:
        data["PO Number"] = match.group(1).strip()

    if not data["PO Number"]:
        for line in text.splitlines():
            if re.search(r"\bP[\s.]*[O0]\b", line, re.IGNORECASE):
                numbers = re.findall(r"\d{5,}", line)
                if numbers:
                    data["PO Number"] = numbers[0]
                    break

    # INV Number
    match = re.search(
        r"\bINV[\s:]*([0-9]{4,})",
        text,
        re.IGNORECASE,
    )
    if match:
        data["INV Number"] = match.group(1).strip()

    if not data["INV Number"]:
        for line in text.splitlines():
            if re.search(r"\bINV\b", line, re.IGNORECASE):
                numbers = re.findall(r"\d{4,}", line)
                if numbers:
                    data["INV Number"] = numbers[0]
                    break

    return data


def get_po_and_tracking(pdf_path):
    """
    Return one result dictionary per PDF page.
    """
    if POPPLER_PATH is None:
        raise FileNotFoundError(
            "Poppler was not found. Expected bundled path: "
            f"{_BUNDLED_POPPLER}"
        )

    images = convert_from_path(
        pdf_path,
        poppler_path=POPPLER_PATH,
    )

    results = []

    for index, image in enumerate(images):
        image_path = f"{pdf_path}_page{index}.png"

        try:
            image.save(image_path)
            text = _ocr_image(image_path)
            results.append(extract_fields(text))
        finally:
            try:
                image.close()
            except Exception:
                pass

            try:
                os.remove(image_path)
            except OSError:
                pass

    return results