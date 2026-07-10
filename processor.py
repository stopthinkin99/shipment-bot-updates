"""
processor.py  —  Orchestrates: extract → parse → validate → write → alert.

Every module this depends on (parser, extractor, excel_writer, mailer)
is force-reloaded on EVERY call, in dependency order, so that files
freshly synced from GitHub always take effect immediately — no stale
cached imports, no restart required.
"""

import json
import importlib


def process_label(pdf_path: str):
    print(f"[INFO] Processing: {pdf_path}")

    # ---- force-reload every dependency, leaf modules first ---- #
    import extractor, parser, excel_writer, mailer
    importlib.reload(extractor)
    importlib.reload(parser)
    importlib.reload(excel_writer)
    importlib.reload(mailer)

    data = {
        "date": "", "ship_to": "", "invoice": "", "carrier": "",
        "tracking_number": "", "sheet": "", "remark": "", "shipper": "",
    }

    try:
        text = extractor.extract_text(pdf_path)
        data = parser.parse_label(text)
        print("[DEBUG DATA]"); print(data)

    except Exception as e:
        _fail(pdf_path, data, f"OCR / parsing failed: {e}", mailer, excel_writer)
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
        _fail(pdf_path, data, reason, mailer, excel_writer,
              write_partial=bool(data.get("sheet")))
        return

    try:
        with open("config.json") as f:
            config = json.load(f)
        excel_writer.update_excel(config["excel_path"], data)
        print(f"[SUCCESS] Written to sheet '{data['sheet']}'")

    except Exception as e:
        _fail(pdf_path, data, f"Excel write failed: {e}", mailer, excel_writer)


def _fail(pdf_path, data, reason, mailer_mod, excel_writer_mod, write_partial=False):
    print(f"[WARN] Manual review needed: {reason}")

    if write_partial and data.get("sheet") and data.get("date"):
        try:
            with open("config.json") as f:
                config = json.load(f)
            excel_writer_mod.update_excel(config["excel_path"], data)
            print(f"[INFO] Partial row written to '{data['sheet']}'")
        except Exception as e:
            print(f"[WARN] Partial write failed: {e}")

    mailer_mod.send_alert(pdf_path, data, reason)