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

    # Common OCR substitutions for UPS prefix.
    if candidate.startswith(("IZ", "I2", "12", "LZ")):
        candidate = "1Z" + candidate[2:]

    # UPS numbers are always 18 characters: 1Z + 16 alphanumerics.
    match = re.search(r"1Z[A-Z0-9]{16}", candidate)
    return match.group(0) if match else ""


def extract_tracking_number(text):
    """Extract UPS, FedEx, USPS, or other long carrier tracking numbers."""
    source = str(text or "")

    # UPS: prioritize text printed after TRACKING / TRACKING #.
    for match in re.finditer(
        r"TRACKING(?:\s+NO\.?|\s+NUMBER)?\s*#?\s*[:\-]?\s*"
        r"([1IL][Z2][A-Z0-9\s\-]{12,45})",
        source,
        re.IGNORECASE,
    ):
        tracking = _normalize_ups_candidate(match.group(1))
        if tracking:
            return tracking

    # UPS fallback anywhere in OCR text, allowing spaces between every group.
    for match in re.finditer(
        r"\b([1IL][Z2](?:[\s\-]*[A-Z0-9]){16})\b",
        source,
        re.IGNORECASE,
    ):
        tracking = _normalize_ups_candidate(match.group(1))
        if tracking:
            return tracking

    compact = re.sub(r"[^A-Z0-9]", "", source.upper())
    for bad_prefix in ("IZ", "I2", "12", "LZ"):
        compact = compact.replace(bad_prefix, "1Z")
    match = re.search(r"1Z[A-Z0-9]{16}", compact)
    if match:
        return match.group(0)

    # FedEx commonly appears as 12 digits in 4-4-4 groups.
    for match in re.finditer(r"\b(\d{4})\s+(\d{4})\s+(\d{4})\b", source):
        return "".join(match.groups())

    # USPS commonly has 20-22 digits and is printed in spaced groups.
    for line in source.splitlines():
        if re.search(r"USPS|TRACKING", line, re.IGNORECASE):
            digits = re.sub(r"\D", "", line)
            if 20 <= len(digits) <= 34:
                return digits

    digits_only = re.sub(r"(?<=\d)[\s\-]+(?=\d)", "", source)
    candidates = re.findall(r"\b\d{10,34}\b", digits_only)

    # Prefer common FedEx/USPS/Brinks length ranges, but avoid ZIP codes etc.
    preferred_lengths = (12, 15, 20, 22, 11, 10)
    for length in preferred_lengths:
        for candidate in candidates:
            if len(candidate) == length:
                return candidate

    return candidates[0] if candidates else ""

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
