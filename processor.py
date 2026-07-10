"""
processor.py  —  Orchestrates: extract → parse → validate → write → alert.
"""

import json
import os
from extractor import extract_text
from parser import parse_label
from excel_writer import update_excel
from mailer import send_alert


def process_label(pdf_path: str):
    print(f"[INFO] Processing: {pdf_path}")

    data = {
        "date": "", "ship_to": "", "invoice": "", "carrier": "",
        "tracking_number": "", "sheet": "", "remark": "", "shipper": "",
    }

    try:
        text = extract_text(pdf_path)
        data = parse_label(text)
        print("[DEBUG DATA]"); print(data)

    except Exception as e:
        _fail(pdf_path, data,
              f"OCR / parsing failed: {e}",
              write_partial=False)
        return

    # ------------------------------------------------------------------ #
    #  VALIDATION — decide what we're confident about
    # ------------------------------------------------------------------ #
    missing = []
    if not data.get("sheet"):
        missing.append("target sheet (invoice prefix not recognised)")
    if not data.get("date"):
        missing.append("ship date")
    if not data.get("tracking_number"):
        missing.append("tracking number")
    if not data.get("ship_to"):
        missing.append("ship-to name")

    if missing:
        reason = "Could not extract: " + ", ".join(missing)
        _fail(pdf_path, data, reason, write_partial=bool(data.get("sheet")))
        return

    # ------------------------------------------------------------------ #
    #  WRITE
    # ------------------------------------------------------------------ #
    try:
        with open("config.json") as f:
            config = json.load(f)

        update_excel(config["excel_path"], data)
        print(f"[SUCCESS] Written to sheet '{data['sheet']}'")

    except Exception as e:
        _fail(pdf_path, data, f"Excel write failed: {e}", write_partial=False)


def _fail(pdf_path, data, reason, write_partial=False):
    """
    Log the failure, optionally write whatever we know to Excel
    (leaving blanks for unknowns), then send an Outlook alert.
    """
    print(f"[WARN] Manual review needed: {reason}")

    if write_partial and data.get("sheet") and data.get("date"):
        try:
            with open("config.json") as f:
                config = json.load(f)
            # mark unknown fields clearly so the reviewer sees gaps
            partial = dict(data)
            print(f"[INFO] Writing partial row to sheet '{data['sheet']}'")
            update_excel(config["excel_path"], partial)
        except Exception as e:
            print(f"[WARN] Partial write also failed: {e}")

    send_alert(pdf_path, data, reason)