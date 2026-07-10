"""
processor.py  —  Orchestrates: extract → parse → validate → write → alert.

Every dependency (parser, extractor, excel_writer, mailer) is
force-reloaded on every call so files freshly synced from GitHub
always take effect immediately.

Writing behaviour: as long as we know which SHEET the shipment
belongs to, we write a row — any field OCR couldn't find gets the
literal text "FILL" so it's visually obvious in Excel what needs a
manual look, instead of silently leaving a blank cell. Only when we
can't even determine the sheet do we skip the Excel write entirely
(nowhere to put the row) and rely on the alert email.
"""

import json
import importlib

FILL = "FILL"   # placeholder written into Excel for any missing field


def process_label(pdf_path: str) -> dict:
    """Returns the extracted data dict (post-fill) for logging by the caller."""
    print(f"[INFO] Processing: {pdf_path}")

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
        print(f"[WARN] OCR / parsing failed: {e}")
        mailer.send_alert(pdf_path, data, f"OCR / parsing failed: {e}")
        return data

    missing = [k for k in ("date", "ship_to", "invoice", "carrier",
                           "tracking_number") if not data.get(k)]

    if not data.get("sheet"):
        reason = "Could not determine target sheet (no matching invoice prefix)"
        print(f"[WARN] {reason}")
        mailer.send_alert(pdf_path, data, reason)
        return data

    # Build the row to write — fill any missing field with "FILL"
    row = dict(data)
    for k in ("date", "ship_to", "invoice", "carrier", "tracking_number"):
        if not row.get(k):
            row[k] = FILL

    try:
        with open("config.json") as f:
            config = json.load(f)
        excel_writer.update_excel(config["excel_path"], row)
        print(f"[SUCCESS] Written to sheet '{data['sheet']}'"
              + (f" (missing: {', '.join(missing)} marked FILL)" if missing else ""))

    except Exception as e:
        print(f"[WARN] Excel write failed: {e}")
        mailer.send_alert(pdf_path, data, f"Excel write failed: {e}")
        return data

    # Still alert if anything was missing, so it gets manually checked
    if missing:
        mailer.send_alert(
            pdf_path, data,
            f"Row written but these fields need manual check: {', '.join(missing)}"
        )

    return row