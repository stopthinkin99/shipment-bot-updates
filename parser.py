import re
import os
import sys
from pathlib import Path
import fitz
import cv2
import numpy as np
import pytesseract
from extractor_core import extract_data_from_file
from po_tracking import get_po_and_tracking
from zales_extractor import extract_zales_from_file, is_zales_label
from malka_brinx import extract_malca_brinks_from_file, is_malca_or_brinks_label

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

SENDER_TO_SHEET = [
    (r"\bEMBY\s+INTERNATIONAL\b", "EMBY"),
    (r"\bEMBY\b", "EMBY"),
    (r"\bFENIX(?:\s+DIAMOND(?:S)?)?\b", "FENIX"),
    (r"\bFEN[1IL]X\b", "FENIX"),
    (r"\bUNI[\s\-]*CREATION\b", "UNI"),
    (r"\bUNI[\s\-]*DESIGN(?:\s+USA)?\b", "UNI"),
    (r"\bUNIVERSAL\s+(?:CREATION|DESIGN)\b", "UNI"),
    (r"\bUNI\b", "UNI"),
]


def _top_label_text(text, max_lines=25):
    lines = [
        line.strip()
        for line in str(text or "").splitlines()
        if line.strip()
    ]
    return "\n".join(lines[:max_lines])


def _route_sheet_by_sender(text):
    header = _top_label_text(text).upper()
    for pattern, sheet in SENDER_TO_SHEET:
        if re.search(pattern, header, re.IGNORECASE):
            return sheet
    return ""


def _quick_ocr(file_path):
    """First-page OCR at 0, 90, and 270 degrees for type and sender detection."""
    doc = fitz.open(file_path)
    try:
        page = doc.load_page(0)
        pix = page.get_pixmap(dpi=250)
        img = np.frombuffer(
            pix.samples,
            dtype=np.uint8,
        ).reshape(pix.h, pix.w, pix.n)
    finally:
        doc.close()

    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    elif pix.n == 3:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

    texts = []
    for angle in (0, 90, 270):
        if angle == 90:
            rotated = cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
        elif angle == 270:
            rotated = cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        else:
            rotated = img

        gray = cv2.cvtColor(rotated, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(
            gray,
            0,
            255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU,
        )
        texts.append(
            pytesseract.image_to_string(binary, config="--oem 3 --psm 6")
        )
        texts.append(
            pytesseract.image_to_string(binary, config="--oem 3 --psm 11")
        )

    return "\n".join(texts)

def parse_label(file_path):
    preview = _quick_ocr(file_path)
    sender_sheet = _route_sheet_by_sender(preview)

    print(f"[DEBUG] Preview text snippet: {preview[:500]}")
    print(f"[INFO] Sender-based sheet: {sender_sheet or 'NOT FOUND'}")

    # ── Zales / UPS ───────────────────────────────────────────────
    if is_zales_label(preview):
        print("[INFO] Zales label detected")
        raw_records = extract_zales_from_file(file_path)
        results = []
        for r in raw_records:
            results.append({
                "sheet":           sender_sheet or r.get("Sheet", "") or "FENIX",
                "date":            r.get("Ship Date", ""),
                "ship_to":         r.get("Recipient Company", ""),
                "invoice":         r.get("PO Number", "") or r.get("INV Number", "") or r.get("Reference", ""),
                "carrier":         "UPS GROUND",
                "tracking_number": r.get("Tracking Number", ""),
                "remark":          "Zales Account",
            })
        return results

    # ── Malca-Amit / Brinks ───────────────────────────────────────
    elif is_malca_or_brinks_label(preview):
        print("[INFO] Malca/Brinks label detected")
        raw_records = extract_malca_brinks_from_file(file_path)
        results = []
        for r in raw_records:
            # malka_brinx.py returns lowercase keys — use them directly
            results.append({
                "sheet":           sender_sheet or r.get("sheet", ""),
                "date":            r.get("date", ""),
                "ship_to":         r.get("ship_to", ""),
                "invoice":         r.get("invoice", ""),
                "carrier":         r.get("carrier", ""),
                "tracking_number": r.get("tracking_number", ""),
                "remark":          r.get("remark", ""),
            })
        return results

    # ── Standard FedEx / UPS flow ─────────────────────────────────
    else:
        print("[INFO] Standard label detected")
        core_records   = extract_data_from_file(file_path)
        po_trk_records = get_po_and_tracking(file_path)

        results = []
        for i, r in enumerate(core_records):
            if i < len(po_trk_records):
                r["Tracking Number"] = po_trk_records[i]["Tracking Number"]
                r["PO Number"]       = po_trk_records[i]["PO Number"]
                if not r.get("INV Number"):
                    r["INV Number"]  = po_trk_records[i].get("INV Number", "")

            invoice = r.get("PO Number") or r.get("INV Number") or r.get("Reference", "")
            carrier = _detect_carrier(r.get("Full Extracted Text", ""))

            results.append({
                "sheet":           sender_sheet or _route_sheet_by_sender(
                    r.get("Full Extracted Text", "")
                ),
                "date":            r.get("Ship Date", ""),
                "ship_to":         r.get("Recipient Company", ""),
                "invoice":         invoice,
                "carrier":         carrier,
                "tracking_number": r.get("Tracking Number", ""),
                "remark":          "",
            })

        return results


def _detect_carrier(text):
    t = text.upper()
    if "UPS" in t:
        if "GROUND" in t:    return "UPS GROUND"
        if "OVERNIGHT" in t: return "UPS O/N"
        return "UPS"
    if "FEDEX" in t or "ORIGIN ID" in t:
        if "PRIORITY" in t:  return "BX FX P/O"
        if "STANDARD" in t:  return "BX FX S/O"
        if "OVERNIGHT" in t: return "BX FX O/N"
        return "FEDEX"
    return ""