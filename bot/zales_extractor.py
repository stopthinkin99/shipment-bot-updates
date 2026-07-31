"""
zales_extractor.py
------------------
UPS/Zales extractor for View/Print Label PDFs.

Version: 2026-07-17.3

These PDFs contain a normal portrait instruction page with the actual UPS
label rotated in the lower half. This extractor crops that lower label,
tests both 90-degree rotations, and accepts only an orientation containing
valid UPS label fields.
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


EXTRACTOR_VERSION = "2026-07-17.4"

import os
import sys
from pathlib import Path

import pytesseract


def _find_tesseract_paths():
    """
    Find bundled or installed Tesseract without using a user-specific path.
    """
    if getattr(sys, "frozen", False):
        app_dir = Path(sys.executable).resolve().parent
    else:
        app_dir = Path(__file__).resolve().parent

    meipass = Path(
        getattr(sys, "_MEIPASS", app_dir)
    )

    candidates = [
        app_dir / "Tesseract-OCR",
        app_dir / "bundle" / "Tesseract-OCR",
        meipass / "Tesseract-OCR",
        Path("C:/Program Files/Tesseract-OCR"),
        Path("C:/Program Files (x86)/Tesseract-OCR"),
    ]

    for folder in candidates:
        exe = folder / "tesseract.exe"
        tessdata = folder / "tessdata"

        if exe.is_file() and tessdata.is_dir():
            return exe, tessdata

    checked = "\n".join(
        f" - {folder}" for folder in candidates
    )

    raise FileNotFoundError(
        "Tesseract OCR could not be found.\n"
        "Checked:\n"
        f"{checked}"
    )


_TESSERACT, _TESSDATA = _find_tesseract_paths()

pytesseract.pytesseract.tesseract_cmd = str(
    _TESSERACT
)

os.environ["TESSDATA_PREFIX"] = str(
    _TESSDATA
)


def _pdf_to_images(file_path):
    document = fitz.open(file_path)
    images = []

    try:
        for page_number in range(len(document)):
            page = document.load_page(page_number)
            pixmap = page.get_pixmap(dpi=300)

            image = np.frombuffer(
                pixmap.samples,
                dtype=np.uint8,
            ).reshape(
                pixmap.h,
                pixmap.w,
                pixmap.n,
            )

            if pixmap.n == 4:
                image = cv2.cvtColor(
                    image,
                    cv2.COLOR_RGBA2BGR,
                )
            elif pixmap.n == 3:
                image = cv2.cvtColor(
                    image,
                    cv2.COLOR_RGB2BGR,
                )

            images.append(image)
    finally:
        document.close()

    return images


def _prepare(image, scale=1.5):
    gray = (
        cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        if len(image.shape) == 3
        else image
    )

    gray = cv2.resize(
        gray,
        None,
        fx=scale,
        fy=scale,
        interpolation=cv2.INTER_CUBIC,
    )

    gray = cv2.GaussianBlur(gray, (3, 3), 0)

    _, binary = cv2.threshold(
        gray,
        0,
        255,
        cv2.THRESH_BINARY + cv2.THRESH_OTSU,
    )

    return binary


def _ocr(image, psm=6, scale=1.5):
    return pytesseract.image_to_string(
        _prepare(image, scale),
        config=f"--oem 3 --psm {psm}",
    )


def is_zales_label(text):
    return bool(
        re.search(
            r"\bZALES(?:\s+BRAND)?\b",
            str(text or ""),
            re.IGNORECASE,
        )
    )


def _normalize_tracking(value):
    candidate = re.sub(
        r"[^A-Z0-9]",
        "",
        str(value or "").upper(),
    )

    if candidate.startswith(("IZ", "I2", "12", "LZ")):
        candidate = "1Z" + candidate[2:]

    if re.fullmatch(r"1Z[A-Z0-9]{16}", candidate):
        return candidate

    return ""


def _extract_tracking(text):
    """
    Extract a UPS number from OCR text.

    This accepts the normal spaced form:
        1Z Y99 G35 29 3066 3817

    It also tolerates OCR splitting TRACKING and the number across adjacent
    lines, but only inside text supplied by the selected UPS label/crop.
    """
    source = str(text or "")

    # First try individual TRACKING lines.
    lines = [
        line.strip()
        for line in source.splitlines()
        if line.strip()
    ]

    for index, line in enumerate(lines):
        if not re.search(r"\bTRACKING\b", line, re.IGNORECASE):
            continue

        # Include the next line in case OCR split the number.
        candidate_text = line
        if index + 1 < len(lines):
            candidate_text += " " + lines[index + 1]

        compact = re.sub(
            r"[^A-Z0-9]",
            "",
            candidate_text.upper(),
        )

        # Find the 18-character UPS value after TRACKING.
        for prefix in ("1Z", "IZ", "I2", "12", "LZ"):
            position = compact.find(prefix)
            if position == -1:
                continue

            candidate = compact[position:position + 18]
            tracking = _normalize_tracking(candidate)

            if tracking:
                return tracking

    # Safe fallback within the selected label/crop only.
    compact = re.sub(r"[^A-Z0-9]", "", source.upper())

    for prefix in ("1Z", "IZ", "I2", "12", "LZ"):
        start_at = 0

        while True:
            position = compact.find(prefix, start_at)
            if position == -1:
                break

            candidate = compact[position:position + 18]
            tracking = _normalize_tracking(candidate)

            if tracking:
                return tracking

            start_at = position + 2

    return ""


def _extract_tracking_from_crop(label_image):
    """
    OCR only the narrow UPS service/tracking band.

    On an upright UPS label this band sits directly below the routing box and
    directly above the large linear barcode. Multiple crop ranges are tried
    because browser print scaling can move the label slightly.
    """
    height, width = label_image.shape[:2]

    crop_ranges = (
        # Main expected tracking band.
        (0.43, 0.61, 0.00, 0.88),
        # Slightly taller fallback.
        (0.39, 0.66, 0.00, 0.92),
        # Whole middle band fallback.
        (0.32, 0.70, 0.00, 0.95),
    )

    for top, bottom, left, right in crop_ranges:
        crop = label_image[
            int(height * top):int(height * bottom),
            int(width * left):int(width * right),
        ]

        if crop.size == 0:
            continue

        for psm in (6, 7, 11, 12):
            crop_text = _ocr(
                crop,
                psm=psm,
                scale=2.5,
            )

            tracking = _extract_tracking(crop_text)

            print(
                f"[ZALES TRACKING CROP] "
                f"range=({top:.2f},{bottom:.2f}) "
                f"psm={psm} tracking={tracking!r}"
            )

            if tracking:
                return tracking

    return ""

def _extract_reference(text):
    match = re.search(
        r"Reference\s*#?\s*1\s*[:\-]?\s*([0-9]{4,})",
        str(text or ""),
        re.IGNORECASE,
    )
    return match.group(1) if match else ""


def _clean_line(value):
    value = re.sub(
        r"[^A-Za-z0-9&.'\-\s]",
        " ",
        str(value or ""),
    )
    return re.sub(r"\s+", " ", value).strip(" .,:;-")


def _valid_recipient(candidate):
    candidate = _clean_line(candidate)

    if not candidate:
        return False

    # A recipient must contain only normal name/company characters.
    if not re.fullmatch(
        r"[A-Za-z][A-Za-z0-9&.'\-]*(?:\s+[A-Za-z0-9&.'\-]+){1,7}",
        candidate,
    ):
        return False

    upper = candidate.upper()

    reject_terms = (
        "SHIP TO",
        "UPS",
        "GROUND",
        "TRACKING",
        "BILLING",
        "REFERENCE",
        "SIGNATURE",
        "REQUIRED",
        "NEW YORK",
        "FIFTH AVENUE",
        "ZALES",
    )

    if any(term in upper for term in reject_terms):
        return False

    if re.search(
        r"\b(?:ST|STREET|AVE|AVENUE|RD|ROAD|BLVD|"
        r"COVE|DR|DRIVE|LANE|LN|COURT|CT)\b",
        upper,
    ):
        return False

    letters = len(re.findall(r"[A-Za-z]", candidate))
    digits = len(re.findall(r"\d", candidate))

    return letters >= 4 and digits <= 1


def _extract_recipient(text):
    lines = [
        _clean_line(line)
        for line in str(text or "").splitlines()
        if _clean_line(line)
    ]

    for index, line in enumerate(lines):
        if not re.search(r"\bSHIP\s*TO\b", line, re.IGNORECASE):
            continue

        for candidate in lines[index + 1:index + 9]:
            digits_only = re.sub(r"\D", "", candidate)

            # UPS account/customer number immediately below SHIP TO.
            if (
                7 <= len(digits_only) <= 15
                and not re.search(r"[A-Za-z]", candidate)
            ):
                continue

            if re.search(
                r"\b[A-Z]{2}\s+\d{5}(?:-\d{4})?\b",
                candidate.upper(),
            ):
                continue

            if _valid_recipient(candidate):
                return candidate

    return ""


def _candidate_score(text):
    recipient = _extract_recipient(text)
    tracking = _extract_tracking(text)
    reference = _extract_reference(text)

    score = 0

    if "SHIP TO" in text.upper():
        score += 10
    if recipient:
        score += 25
    if tracking:
        score += 25
    if reference:
        score += 15
    if "UPS" in text.upper():
        score += 5
    if "ZALES" in text.upper():
        score += 5

    return score, recipient, tracking, reference


def _select_label(page_image):
    height, width = page_image.shape[:2]

    # Actual label occupies the lower portion of UPS View/Print pages.
    crops = [
        page_image[
            int(height * 0.40):int(height * 0.95),
            int(width * 0.01):int(width * 0.99),
        ],
        page_image[
            int(height * 0.44):int(height * 0.92),
            int(width * 0.02):int(width * 0.98),
        ],
    ]

    best = None

    for crop_index, crop in enumerate(crops):
        rotations = (
            (
                "clockwise",
                cv2.rotate(crop, cv2.ROTATE_90_CLOCKWISE),
            ),
            (
                "counterclockwise",
                cv2.rotate(crop, cv2.ROTATE_90_COUNTERCLOCKWISE),
            ),
        )

        for rotation_name, rotated in rotations:
            text = "\n".join(
                [
                    _ocr(rotated, psm=6, scale=1.5),
                    _ocr(rotated, psm=11, scale=1.5),
                ]
            )

            score, recipient, tracking, reference = _candidate_score(text)

            print(
                f"[ZALES ORIENTATION] crop={crop_index} "
                f"rotation={rotation_name} score={score} "
                f"recipient={recipient!r} tracking={tracking!r} "
                f"reference={reference!r}"
            )

            candidate = {
                "score": score,
                "image": rotated,
                "text": text,
                "recipient": recipient,
                "tracking": tracking,
                "reference": reference,
            }

            if best is None or candidate["score"] > best["score"]:
                best = candidate

    if best is None:
        raise RuntimeError("No UPS label orientation could be evaluated.")

    return best


def parse_zales_label(
    text,
    label_image,
    filename,
    page,
    *,
    recipient="",
    tracking="",
    reference="",
):
    return {
        "Source File": filename,
        "Page": page,
        "Sheet": "FENIX",
        "Tracking Number": (
            tracking
            or _extract_tracking_from_crop(label_image)
            or _extract_tracking(text)
        ),
        "PO Number": reference or _extract_reference(text),
        "Recipient Company": recipient or _extract_recipient(text),
        "INV Number": "",
        "Reference": "",
        "CAD": "",
        "Weight": "",
        "Ship Date": datetime.now().strftime("%Y-%m-%d"),
        "Full Extracted Text": text,
        "Application Run Date and time": (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ),
    }


def extract_zales_from_file(file_path):
    print(
        f"[ZALES] ZALES EXTRACTOR VERSION: {EXTRACTOR_VERSION}"
    )

    filename = os.path.basename(file_path)
    records = []

    for page_index, page_image in enumerate(
        _pdf_to_images(file_path),
        start=1,
    ):
        selected = _select_label(page_image)

        print(
            f"[ZALES SELECTED] page={page_index} "
            f"score={selected['score']} "
            f"recipient={selected['recipient']!r} "
            f"tracking={selected['tracking']!r} "
            f"reference={selected['reference']!r}"
        )

        # A Zales label is established by either visible Zales text or the
        # combination of a valid UPS tracking number and Reference #1.
        if (
            is_zales_label(selected["text"])
            or (
                selected["tracking"]
                and selected["reference"]
            )
        ):
            records.append(
                parse_zales_label(
                    selected["text"],
                    selected["image"],
                    filename,
                    page_index,
                    recipient=selected["recipient"],
                    tracking=selected["tracking"],
                    reference=selected["reference"],
                )
            )

    return records