"""
processor.py
------------
Coordinates label extraction, manual correction, Excel writing, and
fallback email alerts.

New behavior:
- Complete records are written directly to Excel.
- Incomplete records open the manual-review popup before Excel is updated.
- Extracted values are prefilled; the user completes only missing fields.
- The corrected record is written after the user clicks Done.
- A cancelled/failed review is not written to Excel.
- Email alerts remain only as a fallback for extraction failure, popup
  cancellation/failure, or Excel-write failure.
"""

import os
import shutil
from datetime import date
from typing import Callable, Iterable, Optional

from parser import parse_label
from excel_writer import update_excel
from mailer import send_alert


# Date is automatically filled with today's date when OCR misses it.
# These are the fields the popup asks the user to verify/complete.
REQUIRED_REVIEW_FIELDS = (
    "sheet",
    "ship_to",
    "invoice",
    "tracking_number",
    "carrier",
)

FIELD_LABELS = {
    "sheet": "Excel Sheet",
    "ship_to": "Ship To",
    "invoice": "Invoice / PO",
    "tracking_number": "Tracking Number",
    "carrier": "Carrier",
}


ReviewCallback = Callable[[str, dict], Optional[dict]]


def _is_blank(value) -> bool:
    return value is None or not str(value).strip()


def _missing_fields(record: dict) -> list[str]:
    """Return popup-required record keys whose values are blank."""
    return [
        field
        for field in REQUIRED_REVIEW_FIELDS
        if _is_blank(record.get(field))
    ]


def _format_missing_fields(fields: Iterable[str]) -> str:
    labels = [FIELD_LABELS.get(field, field) for field in fields]
    return ", ".join(labels)


def _prepare_record(record: dict) -> dict:
    """
    Normalize a parsed record before review.

    The popup requested by the shipping team does not contain a date field,
    so a missing date is safely defaulted to today's ISO date.
    """
    prepared = dict(record)

    if _is_blank(prepared.get("date")):
        prepared["date"] = date.today().strftime("%Y-%m-%d")

    # Ensure all popup fields exist so the dialog can prefill them.
    for field in REQUIRED_REVIEW_FIELDS:
        prepared.setdefault(field, "")

    prepared.setdefault("remark", "")
    return prepared


def _move_to_done(
    file_path: str,
    done_folder: str,
    log_fn=print,
) -> None:
    os.makedirs(done_folder, exist_ok=True)

    destination = os.path.join(
        done_folder,
        os.path.basename(file_path),
    )

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


def process_file(
    file_path,
    excel_path=None,
    done_folder=None,
    log_fn=print,
    review_callback: ReviewCallback | None = None,
):
    """
    Extract `file_path`, request manual correction when necessary, and write
    completed records to `excel_path`.

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

        for index, raw_record in enumerate(records, start=1):
            if not isinstance(raw_record, dict):
                reason = (
                    f"Extractor returned an invalid record on page {index}: "
                    f"{type(raw_record).__name__}"
                )
                log_fn(f"[ERROR] {reason}")

                send_alert(
                    pdf_path=file_path,
                    data={},
                    reason=reason,
                    log_fn=log_fn,
                )
                continue

            record = _prepare_record(raw_record)
            missing = _missing_fields(record)

            if missing:
                missing_text = _format_missing_fields(missing)
                log_fn(
                    "[REVIEW] Required fields need completion before "
                    f"Excel is updated: {missing_text}."
                )

                if review_callback is None:
                    reason = (
                        "Manual review was required, but app.py did not "
                        "provide a review callback. Missing: "
                        f"{missing_text}."
                    )
                    log_fn(f"[REVIEW ERROR] {reason}")

                    send_alert(
                        pdf_path=file_path,
                        data=record,
                        reason=reason,
                        log_fn=log_fn,
                    )
                    continue

                try:
                    corrected = review_callback(
                        file_path,
                        record,
                    )
                except Exception as exc:
                    reason = (
                        "Manual-review popup failed: "
                        f"{type(exc).__name__}: {exc}"
                    )
                    log_fn(f"[REVIEW ERROR] {reason}")

                    send_alert(
                        pdf_path=file_path,
                        data=record,
                        reason=reason,
                        log_fn=log_fn,
                    )
                    continue

                if corrected is None:
                    reason = (
                        "The user cancelled manual review. "
                        "The shipment was not written to Excel."
                    )
                    log_fn(f"[REVIEW CANCELLED] {reason}")

                    send_alert(
                        pdf_path=file_path,
                        data=record,
                        reason=reason,
                        log_fn=log_fn,
                    )
                    continue

                if not isinstance(corrected, dict):
                    reason = (
                        "Manual-review popup returned an invalid result: "
                        f"{type(corrected).__name__}"
                    )
                    log_fn(f"[REVIEW ERROR] {reason}")

                    send_alert(
                        pdf_path=file_path,
                        data=record,
                        reason=reason,
                        log_fn=log_fn,
                    )
                    continue

                record = _prepare_record(corrected)
                remaining = _missing_fields(record)

                if remaining:
                    remaining_text = _format_missing_fields(remaining)
                    reason = (
                        "Manual review closed with required fields still "
                        f"missing: {remaining_text}."
                    )
                    log_fn(f"[REVIEW ERROR] {reason}")

                    send_alert(
                        pdf_path=file_path,
                        data=record,
                        reason=reason,
                        log_fn=log_fn,
                    )
                    continue

                log_fn(
                    "[REVIEW] Manual completion accepted. "
                    "Writing corrected shipment to Excel."
                )

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

        if not written_records:
            log_fn("[ERROR] Nothing was written to Excel.")
            return []

        if done_folder:
            _move_to_done(
                file_path,
                done_folder,
                log_fn,
            )

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