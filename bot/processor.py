import json
import os
import importlib
from extractor import extract_text
from parser import parse_label
from excel_writer import update_excel


def process_label(pdf_path: str):
    print(f"[INFO] Processing: {pdf_path}")

    # Reload mailer fresh each call — picks up any file changes immediately
    import mailer
    importlib.reload(mailer)

    data = {
        "date": "", "ship_to": "", "invoice": "", "carrier": "",
        "tracking_number": "", "sheet": "", "remark": "", "shipper": "",
    }

    try:
        text = extract_text(pdf_path)
        data = parse_label(text)
        print("[DEBUG DATA]"); print(data)

    except Exception as e:
        _fail(pdf_path, data, f"OCR / parsing failed: {e}", mailer)
        return

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
        _fail(pdf_path, data, reason,
              mailer, write_partial=bool(data.get("sheet")))
        return

    try:
        with open("config.json") as f:
            config = json.load(f)
        update_excel(config["excel_path"], data)
        print(f"[SUCCESS] Written to sheet '{data['sheet']}'")

    except Exception as e:
        _fail(pdf_path, data, f"Excel write failed: {e}", mailer)


def _fail(pdf_path, data, reason, mailer_mod, write_partial=False):
    print(f"[WARN] Manual review needed: {reason}")

    if write_partial and data.get("sheet") and data.get("date"):
        try:
            with open("config.json") as f:
                config = json.load(f)
            update_excel(config["excel_path"], data)
            print(f"[INFO] Partial row written to '{data['sheet']}'")
        except Exception as e:
            print(f"[WARN] Partial write failed: {e}")

    mailer_mod.send_alert(pdf_path, data, reason)