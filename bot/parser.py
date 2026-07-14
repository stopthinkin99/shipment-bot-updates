import re
from extractor_core import extract_data_from_file
from po_tracking import get_po_and_tracking
from zales_extractor import extract_zales_from_file, is_zales_label
import fitz
import cv2
import numpy as np
import pytesseract
import os
from malca_amit import extract_malca_brinks_from_file, is_malca_or_brinks_label
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tessdata"

# ── Invoice prefix → sheet routing (unchanged from original) ──────────────
PREFIX_SHEET = [
    ("2030", "EMBY"),
    ("82",   "FENIX"),
    ("47",   "FENIX"),
    ("10",   "UNI"),
]

def _route_sheet(invoice_number):
    for prefix, sheet in PREFIX_SHEET:
        if str(invoice_number).startswith(prefix):
            return sheet
    return ""

def _quick_ocr(file_path):
    """Low-res first-page OCR just for Zales detection."""
    doc = fitz.open(file_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=150)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    doc.close()
    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    return pytesseract.image_to_string(img)

def parse_label(file_path):
    """
    Entry point called by processor.py.
    Returns a list of dicts (one per page) with keys:
      sheet, date, ship_to, invoice, carrier, tracking_number, remark
    """
    preview = _quick_ocr(file_path)

    if is_zales_label(preview):
        raw_records = extract_zales_from_file(file_path)
        results = []
        for r in raw_records:
            results.append({
                "sheet":          "FENIX",
                "date":           r.get("Ship Date", ""),
                "ship_to":        r.get("Recipient Company", ""),
                "invoice":        r.get("PO Number", "") or r.get("INV Number", "") or r.get("Reference", ""),
                "carrier":        "UPS GROUND",
                "tracking_number": r.get("Tracking Number", ""),
                "remark":         "Zales Account",
            })
        return results

    elif is_malca_or_brinks_label(preview):
        print(["[INFO] Malca/Brinks label detected"])
        raw_records = extract_malca_brinks_from_file(file_path)
        results = []
        for r in raw_records:
            results.append({
                "sheet":          r.get("Sheet", ""),
                "date":           r.get("date", ""),
                "ship_to":        r.get("ship_to", ""),
                "invoice":        r.get("PO Number", "") or r.get("INV Number", "") or r.get("Reference", ""),
                "carrier":        r.get("Carrier", ""),
                "tracking_number": r.get("Tracking Number", ""),
                "remark":         "",
            })
        return results

    # Standard flow
    core_records   = extract_data_from_file(file_path)
    po_trk_records = get_po_and_tracking(file_path)

    results = []
    for i, r in enumerate(core_records):
        # Merge PO/tracking from po_tracking.py
        if i < len(po_trk_records):
            r["Tracking Number"] = po_trk_records[i]["Tracking Number"]
            r["PO Number"]       = po_trk_records[i]["PO Number"]
            # Only fill INV from po_tracking if core didn't already find one
            if not r.get("INV Number"):
                r["INV Number"] = po_trk_records[i].get("INV Number", "")

        # Pick best invoice field: PO first, then INV, then REF
        invoice = r.get("PO Number") or r.get("INV Number") or r.get("Reference", "")

        # Determine carrier from OCR text
        raw_text = r.get("Full Extracted Text", "")
        carrier = _detect_carrier(raw_text)

        results.append({
            "sheet":           _route_sheet(invoice),
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
        if "GROUND" in t:     return "UPS GROUND"
        if "OVERNIGHT" in t:  return "UPS O/N"
        return "UPS"
    if "FEDEX" in t or "ORIGIN ID" in t:
        if "PRIORITY" in t:   return "BX FX P/O"
        if "STANDARD" in t:   return "BX FX S/O"
        if "OVERNIGHT" in t:  return "BX FX O/N"
        return "FEDEX"
    return ""