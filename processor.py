"""
processor.py
------------
Coordinates label extraction, Excel writing, and manual-review email alerts.

Behavior:
- If no records are extracted: log an error and return [].
- If the target sheet is missing: send a manual-review alert and skip the row.
- If other required fields are missing: write the partial row to Excel and
  send a manual-review alert identifying what must be filled manually.
- Return only records successfully written to Excel.
"""

import os
import shutil
from typing import Iterable

from parser import parse_label
from excel_writer import update_excel
from mailer import send_alert


# Fields expected for a complete shipment row.
# Remark is optional and is therefore not included.
REQUIRED_FIELDS = (
    "date",
    "ship_to",
    "invoice",
    "carrier",
    "tracking_number",
    "sheet",
)

FIELD_LABELS = {
    "date": "Date",
    "ship_to": "Ship To",
    "invoice": "Invoice / PO",
    "carrier": "Carrier",
    "tracking_number": "Tracking Number",
    "sheet": "Excel Sheet",
}


def _is_blank(value) -> bool:
    return value is None or not str(value).strip()


def _missing_fields(record: dict) -> list[str]:
    """Return required record keys whose values are blank."""
    return [
        field
        for field in REQUIRED_FIELDS
        if _is_blank(record.get(field))
    ]


def _format_missing_fields(fields: Iterable[str]) -> str:
    labels = [FIELD_LABELS.get(field, field) for field in fields]
    return ", ".join(labels)


def process_file(
    file_path,
    excel_path=None,
    done_folder=None,
    log_fn=print,
):
    """
    Extract `file_path`, write usable records to `excel_path`, and send
    Microsoft Graph manual-review alerts for incomplete records.

    Returns:
        list[dict]: records successfully written to Excel.
    """
    if not excel_path:
        log_fn("[ERROR] No Excel tracking file was supplied by app.py.")
        return []

    if not os.path.isfile(excel_path):
        log_fn(
            f"[ERROR] Excel tracking file does not exist: {excel_path}"
        )
        return []

    if not os.path.isfile(file_path):
        log_fn(f"[ERROR] Label file does not exist: {file_path}")
        return []

    log_fn(f"[INFO] Processing: {file_path}")
    log_fn(f"[INFO] Excel target: {excel_path}")

    try:
        records = parse_label(file_path)

        if not records:
            reason = "No shipment records could be extracted from the PDF."
            log_fn(f"[ERROR] {reason}")

            send_alert(
                pdf_path=file_path,
                data={},
                reason=reason,
                log_fn=log_fn,
            )
            return []

        written_records = []

        for index, record in enumerate(records, start=1):
            if not isinstance(record, dict):
                reason = (
                    f"Extractor returned an invalid record on page {index}: "
                    f"{type(record).__name__}"
                )
                log_fn(f"[ERROR] {reason}")

                send_alert(
                    pdf_path=file_path,
                    data={},
                    reason=reason,
                    log_fn=log_fn,
                )
                continue

            missing = _missing_fields(record)

            # Without a sheet, Excel cannot know where to write the row.
            if "sheet" in missing:
                invoice = record.get("invoice") or "unknown"
                reason = (
                    "Could not determine the target Excel sheet for "
                    f"invoice '{invoice}'."
                )

                log_fn(f"[SKIP] {reason}")

                send_alert(
                    pdf_path=file_path,
                    data=record,
                    reason=reason,
                    log_fn=log_fn,
                )
                continue

            # Write the row even when other fields are missing. This leaves
            # the partial information in Excel for shipping staff to complete.
            try:
                update_excel(excel_path, record)
                written_records.append(record)

                log_fn(
                    f"[OK] Written to sheet '{record['sheet']}': "
                    f"invoice={record.get('invoice') or '—'}"
                )
            except Exception as exc:
                reason = (
                    "Excel write failed: "
                    f"{type(exc).__name__}: {exc}"
                )
                log_fn(f"[ERROR] {reason}")

                send_alert(
                    pdf_path=file_path,
                    data=record,
                    reason=reason,
                    log_fn=log_fn,
                )
                continue

            # The row was written, but shipping must manually complete it.
            missing_after_sheet = [
                field
                for field in missing
                if field != "sheet"
            ]

            if missing_after_sheet:
                missing_text = _format_missing_fields(
                    missing_after_sheet
                )
                reason = (
                    "Shipment was written to Excel, but these required "
                    f"fields are missing: {missing_text}."
                )

                log_fn(f"[REVIEW] {reason}")

                send_alert(
                    pdf_path=file_path,
                    data=record,
                    reason=reason,
                    log_fn=log_fn,
                )

        if not written_records:
            log_fn("[ERROR] Nothing was written to Excel.")
            return []

        # Moving is optional. If no Done folder is supplied, leave the PDF
        # in the watched directory.
        if done_folder:
            os.makedirs(done_folder, exist_ok=True)
            destination = os.path.join(
                done_folder,
                os.path.basename(file_path),
            )

            # Avoid crashing when destination already exists.
            if os.path.exists(destination):
                base, extension = os.path.splitext(destination)
                counter = 1
                candidate = f"{base}_{counter}{extension}"

                while os.path.exists(candidate):
                    counter += 1
                    candidate = f"{base}_{counter}{extension}"

                destination = candidate

            shutil.move(file_path, destination)
            log_fn(f"[DONE] Moved to {destination}")

        return written_records

    except Exception as exc:
        reason = (
            f"Processing failed: {type(exc).__name__}: {exc}"
        )
        log_fn(f"[ERROR] {reason}")

        send_alert(
            pdf_path=file_path,
            data={},
            reason=reason,
            log_fn=log_fn,
        )
        return []


# Existing app.py imports processor.process_label.
process_label = process_file