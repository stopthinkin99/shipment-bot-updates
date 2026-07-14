import os
import shutil

from parser import parse_label
from excel_writer import update_excel


def process_file(file_path, excel_path=None, done_folder=None, log_fn=print):
    """
    Extract a label, write it to the Excel workbook selected in app.py,
    and optionally move the PDF to a Done folder.

    Returns a list of extracted records on success.
    Returns [] after logging the real error on failure.
    """
    if not excel_path:
        log_fn("[ERROR] No Excel tracking file was supplied by app.py.")
        return []

    if not os.path.isfile(excel_path):
        log_fn(f"[ERROR] Excel tracking file does not exist: {excel_path}")
        return []

    log_fn(f"[INFO] Processing: {file_path}")
    log_fn(f"[INFO] Excel target: {excel_path}")

    try:
        records = parse_label(file_path)

        if not records:
            log_fn(f"[WARN] No records extracted from {file_path}")
            return []

        written_records = []

        for record in records:
            if not record.get("sheet"):
                log_fn(
                    "[SKIP] Could not determine sheet for invoice "
                    f"'{record.get('invoice')}' — manual review needed"
                )
                continue

            if not record.get("tracking_number"):
                log_fn(f"[WARN] No tracking number found in {file_path}")

            update_excel(excel_path, record)
            written_records.append(record)
            log_fn(
                f"[OK] Written to sheet '{record['sheet']}': "
                f"invoice={record.get('invoice') or '—'}"
            )

        if not written_records:
            log_fn("[ERROR] Nothing was written to Excel.")
            return []

        # Moving is optional. The GUI currently does not need to provide
        # a Done folder, so the original PDF can remain in the watched folder.
        if done_folder:
            os.makedirs(done_folder, exist_ok=True)
            destination = os.path.join(done_folder, os.path.basename(file_path))
            shutil.move(file_path, destination)
            log_fn(f"[DONE] Moved to {done_folder}")

        return written_records

    except Exception as exc:
        log_fn(f"[ERROR] Failed on {file_path}: {type(exc).__name__}: {exc}")
        return []


process_label = process_file