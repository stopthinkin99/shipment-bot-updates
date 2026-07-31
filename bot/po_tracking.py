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
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)

_MEIPASS_DIR = Path(getattr(sys, "_MEIPASS", _BASE_DIR))


def _find_poppler():
    candidates = [
        _BASE_DIR / "poppler" / "bin",
        _BASE_DIR / "bundle" / "poppler" / "bin",
        _MEIPASS_DIR / "poppler" / "bin",
        Path("C:/Program Files/poppler/Library/bin"),
        Path("C:/Program Files (x86)/poppler/Library/bin"),
    ]

    for folder in candidates:
        if (folder / "pdfinfo.exe").is_file():
            return folder

    return None


def _find_tesseract():
    candidates = [
        _BASE_DIR / "Tesseract-OCR",
        _BASE_DIR / "bundle" / "Tesseract-OCR",
        _MEIPASS_DIR / "Tesseract-OCR",
        Path("C:/Program Files/Tesseract-OCR"),
        Path("C:/Program Files (x86)/Tesseract-OCR"),
    ]

    for folder in candidates:
        exe = folder / "tesseract.exe"
        tessdata = folder / "tessdata"

        if exe.is_file() and tessdata.is_dir():
            return exe, tessdata

    raise FileNotFoundError(
        "Tesseract OCR could not be found in the bundled or standard install locations."
    )


POPPLER_PATH_OBJ = _find_poppler()
POPPLER_PATH = str(POPPLER_PATH_OBJ) if POPPLER_PATH_OBJ else None

_TESSERACT, _TESSDATA = _find_tesseract()

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



def _normalize_document_number(value):
    value = str(value or "").strip()
    value = re.sub(r"\s+", " ", value)
    value = re.sub(r"\s*/\s*", "/", value)
    value = re.sub(r"\s*-\s*", "-", value)
    value = re.sub(r"\s*,\s*", ", ", value)
    return value.strip(" .,:;-")


def _extract_document_number(text, labels):
    label_pattern = "|".join(labels)

    for raw_line in str(text or "").splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        match = re.search(
            rf"\b(?:{label_pattern})\b"
            r"\s*(?:NUMBER|NO\.?|#)?\s*[:\-]?\s*"
            r"([A-Z0-9][A-Z0-9/,\- ]{2,60})",
            line,
            re.IGNORECASE,
        )
        if not match:
            continue

        candidate = match.group(1)
        candidate = re.split(
            r"\s{2,}|\b(?:DATE|SHIP TO|TRACKING|WEIGHT|PHONE|REF(?:ERENCE)?)\b",
            candidate,
            maxsplit=1,
            flags=re.IGNORECASE,
        )[0]
        candidate = _normalize_document_number(candidate)

        if re.search(r"\d", candidate):
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
        "Tracking Number": extract_tracking_number(text),
        "PO Number": _extract_document_number(
            text,
            (r"P[\s.]*[O0]", r"PURCHASE\s+ORDER"),
        ),
        "INV Number": _extract_document_number(
            text,
            (r"I[\s.]*N[\s.]*V(?:OICE)?", r"INVOICE", r"MEMO"),
        ),
    }
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

