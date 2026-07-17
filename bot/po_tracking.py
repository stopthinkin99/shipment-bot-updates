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


def _normalize_ups_candidate(value):
    candidate = re.sub(r"[^A-Z0-9]", "", str(value or "").upper())

    if candidate.startswith(("IZ", "I2", "12", "LZ")):
        candidate = "1Z" + candidate[2:]

    # UPS numbers must be exactly 18 characters.
    if re.fullmatch(r"1Z[A-Z0-9]{16}", candidate):
        return candidate

    return ""


def extract_tracking_number(text):
    """
    Extract tracking safely.

    Never compact the entire OCR page and search for 1Z because that can
    join unrelated fields into a fake value such as:
        1Z + phone number + SHIP DATE
    """
    source = str(text or "")
    lines = [line.strip() for line in source.splitlines() if line.strip()]

    # UPS: require a real TRACKING line.
    for line in lines:
        if not re.search(
            r"\bTRACKING(?:\s+NO\.?|\s+NUMBER)?\s*#?",
            line,
            re.IGNORECASE,
        ):
            continue

        match = re.search(
            r"([1IL][Z2](?:[\s\-]*[A-Z0-9]){16})",
            line,
            re.IGNORECASE,
        )
        if match:
            tracking = _normalize_ups_candidate(match.group(1))
            if tracking:
                return tracking

    # UPS fallback: bounded standalone 1Z + 16 characters.
    for match in re.finditer(
        r"(?<![A-Z0-9])"
        r"([1IL][Z2](?:[\s\-]*[A-Z0-9]){16})"
        r"(?![A-Z0-9])",
        source,
        re.IGNORECASE,
    ):
        tracking = _normalize_ups_candidate(match.group(1))
        if tracking:
            return tracking

    # FedEx: prioritize the TRK# / TRACKING line.
    for line in lines:
        if not re.search(r"\b(?:TRK|TRACKING)\s*#?", line, re.IGNORECASE):
            continue

        grouped = re.search(
            r"(?<!\d)(\d{4})\s+(\d{4})\s+(\d{4})(?!\d)",
            line,
        )
        if grouped:
            return "".join(grouped.groups())

        candidates = re.findall(r"(?<!\d)(\d{12}|\d{15})(?!\d)", line)
        if candidates:
            return candidates[-1]

    # FedEx generic 4-4-4 pattern.
    for match in re.finditer(
        r"(?<!\d)(\d{4})\s+(\d{4})\s+(\d{4})(?!\d)",
        source,
    ):
        return "".join(match.groups())

    # FedEx continuous 12/15 digits only when the page is clearly FedEx.
    if re.search(
        r"\b(?:FEDEX|FED\s*EX|STANDARD\s+OVERNIGHT|"
        r"PRIORITY\s+OVERNIGHT)\b",
        source,
        re.IGNORECASE,
    ):
        for line in lines:
            match = re.search(r"(?<!\d)(\d{12}|\d{15})(?!\d)", line)
            if match:
                return match.group(1)

    # USPS.
    if re.search(r"\bUSPS\b", source, re.IGNORECASE):
        for line in lines:
            digits = re.sub(r"\D", "", line)
            if len(digits) in {20, 22}:
                return digits

    # Brinks/armored labeled tracking.
    for line in lines:
        if re.search(
            r"\b(?:TRACKING\s*NO|HAWB\s*NUMBER)\b",
            line,
            re.IGNORECASE,
        ):
            digits = re.sub(r"\D", "", line)
            if 8 <= len(digits) <= 15:
                return digits

    return ""

def _ocr_pil_image(image):
    gray = image.convert("L")
    texts = []
    for angle in (0, 90, 270):
        rotated = gray if angle == 0 else gray.rotate(angle, expand=True, fillcolor=255)
        for psm in (6, 11):
            texts.append(
                pytesseract.image_to_string(
                    rotated,
                    config=f"--oem 3 --psm {psm}",
                )
            )
    return "\n".join(texts)

def extract_fields(text):
    data = {
        "Tracking Number": "",
        "PO Number": "",
        "INV Number": "",
    }

    data["Tracking Number"] = extract_tracking_number(text)

    # PO Number, including Brinks abbreviated lists such as:
    # PO # 86100087, 89, 90, 107
    po_line_match = re.search(
        r"\bP[\s.]*[O0]\s*#?\s*[:\-]?\s*"
        r"([0-9]{5,}(?:\s*[,;/]\s*[0-9]{1,8})+|[0-9]{5,})",
        text,
        re.IGNORECASE,
    )
    if po_line_match:
        raw_po = po_line_match.group(1).strip()
        parts = re.findall(r"\d+", raw_po)
        if parts:
            first = parts[0]
            expanded = [first]
            for suffix in parts[1:]:
                if len(suffix) < len(first):
                    expanded.append(first[:-len(suffix)] + suffix)
                else:
                    expanded.append(suffix)
            data["PO Number"] = ", ".join(expanded)

    if not data["PO Number"]:
        for line in text.splitlines():
            if re.search(r"\bP[\s.]*[O0]\b", line, re.IGNORECASE):
                values = re.findall(r"\d{5,}", line)
                if values:
                    data["PO Number"] = values[0]
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
    """Return tracking, PO, and INV data for every PDF page."""
    if POPPLER_PATH is None:
        raise FileNotFoundError(
            "Poppler was not found. Expected bundled folder: "
            f"{_BUNDLED_POPPLER}"
        )

    images = convert_from_path(
        pdf_path,
        dpi=300,
        poppler_path=POPPLER_PATH,
    )

    results = []
    for image in images:
        try:
            text = _ocr_pil_image(image)
            results.append(extract_fields(text))
        finally:
            try:
                image.close()
            except Exception:
                pass

    return results
