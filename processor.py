"""
processor.py
------------
Coordinates label extraction, manual correction, and Excel writing.

New behavior:
- Complete records are written directly to Excel.
- Incomplete records open the manual-review popup before Excel is updated.
- Extracted values are prefilled; the user completes only missing fields.
- The corrected record is written after the user clicks Done.
- A cancelled/failed review is not written to Excel.
- Processing problems are handled through the manual-review popup.
- No shipment-team email is sent by this module.
"""

import os
import re
import shutil
from datetime import date
from typing import Callable, Iterable, Optional

from parser import parse_label
from excel_writer import update_excel


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


def _compact(value) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def _tracking_is_valid(value, carrier) -> bool:
    compact = _compact(value)
    carrier_upper = str(carrier or "").upper().strip()

    if not compact:
        return False

    if compact.startswith("1Z"):
        return bool(
            re.fullmatch(r"1Z[A-Z0-9]{16}", compact)
            and len(re.findall(r"\d", compact)) >= 8
        )

    if "UPS" in carrier_upper:
        return bool(
            re.fullmatch(r"1Z[A-Z0-9]{16}", compact)
            and len(re.findall(r"\d", compact)) >= 8
        )

    if (
        "FEDEX" in carrier_upper
        or "FX" in carrier_upper
        or re.fullmatch(r"M\s*/\s*E", carrier_upper)
    ):
        return bool(re.fullmatch(r"\d{12}|\d{15}", compact))

    if "USPS" in carrier_upper or "POSTAL" in carrier_upper:
        return bool(re.fullmatch(r"\d{20,34}", compact))

    if "MALCA" in carrier_upper or "MALKA" in carrier_upper:
        return bool(re.fullmatch(r"\d{5,20}", compact))

    if "BRINKS" in carrier_upper:
        return bool(re.fullmatch(r"\d{7,20}", compact))

    return (
        7 <= len(compact) <= 34
        and len(re.findall(r"\d", compact)) >= 6
    )


def _invoice_is_valid(value) -> bool:
    raw = re.sub(r"\s+", " ", str(value or "")).strip()

    if not raw:
        return False

    if not re.fullmatch(r"[A-Z0-9][A-Z0-9/,\- ]{2,60}", raw, re.I):
        return False

    numeric_tokens = re.findall(r"\d+", raw)

    if len(numeric_tokens) >= 2:
        first = numeric_tokens[0]

        for token in numeric_tokens[1:]:
            if len(token) > len(first) and token.startswith(first):
                return False

        if (
            not re.search(r"[/,\-]", raw)
            and any(len(token) >= len(first) for token in numeric_tokens[1:])
        ):
            return False

    return True


def _invalidate_suspicious_fields(record: dict, log_fn=print) -> dict:
    checked = dict(record)

    tracking = checked.get("tracking_number", "")
    carrier = checked.get("carrier", "")
    if tracking and not _tracking_is_valid(tracking, carrier):
        log_fn(
            "[REVIEW] Suspicious tracking number rejected: "
            f"{tracking!r}. Manual entry required."
        )
        checked["_rejected_tracking_number"] = tracking
        checked["tracking_number"] = ""

    invoice = checked.get("invoice", "")
    if invoice and not _invoice_is_valid(invoice):
        log_fn(
            "[REVIEW] Suspicious Invoice/PO/Memo value rejected: "
            f"{invoice!r}. Manual entry required."
        )
        checked["_rejected_invoice"] = invoice
        checked["invoice"] = ""

    return checked


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


def _request_manual_review(
    file_path: str,
    record: dict,
    review_callback: ReviewCallback | None,
    log_fn=print,
) -> dict | None:
    """
    Open the existing manual-review popup.

    Returns a complete validated record, or None when review cannot be
    completed. No fallback email is sent.
    """
    prepared = _prepare_record(record)
    prepared = _invalidate_suspicious_fields(prepared, log_fn)
    missing = _missing_fields(prepared)

    if not missing:
        return prepared

    missing_text = _format_missing_fields(missing)
    log_fn(
        "[REVIEW] Required fields need completion before Excel is updated: "
        f"{missing_text}."
    )

    if review_callback is None:
        log_fn(
            "[REVIEW ERROR] Manual review was required, but no review "
            "callback was provided. Shipment was not written."
        )
        return None

    try:
        corrected = review_callback(file_path, prepared)
    except Exception as exc:
        log_fn(
            "[REVIEW ERROR] Manual-review popup failed: "
            f"{type(exc).__name__}: {exc}"
        )
        return None

    if corrected is None:
        log_fn(
            "[REVIEW CANCELLED] Shipment was not written to Excel."
        )
        return None

    if not isinstance(corrected, dict):
        log_fn(
            "[REVIEW ERROR] Manual-review popup returned an invalid result: "
            f"{type(corrected).__name__}"
        )
        return None

    corrected = _prepare_record(corrected)
    corrected = _invalidate_suspicious_fields(corrected, log_fn)
    remaining = _missing_fields(corrected)

    if remaining:
        log_fn(
            "[REVIEW ERROR] Required fields are still missing: "
            f"{_format_missing_fields(remaining)}."
        )
        return None

    log_fn(
        "[REVIEW] Manual completion accepted. "
        "Writing corrected shipment to Excel."
    )
    return corrected


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
    Extract a label, request manual correction whenever extraction or
    validation fails, and write only a complete record to Excel.

    No email alert is sent by this workflow.
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
    except Exception as exc:
        # Parser failure must still open the popup instead of emailing.
        log_fn(
            "[ERROR] Parser failed: "
            f"{type(exc).__name__}: {exc}"
        )
        records = [{}]

    if not records:
        log_fn(
            "[REVIEW] No shipment record was extracted. "
            "Opening manual-review popup."
        )
        records = [{}]

    written_records = []

    for index, raw_record in enumerate(records, start=1):
        if not isinstance(raw_record, dict):
            log_fn(
                "[REVIEW] Extractor returned an invalid record on "
                f"page {index}; opening manual-review popup."
            )
            raw_record = {}

        record = _prepare_record(raw_record)
        record = _invalidate_suspicious_fields(record, log_fn)

        if _missing_fields(record):
            record = _request_manual_review(
                file_path,
                record,
                review_callback,
                log_fn,
            )

            if record is None:
                continue

        try:
            update_excel(excel_path, record)
            written_records.append(record)

            log_fn(
                f"[OK] Written to sheet '{record['sheet']}': "
                f"invoice={record.get('invoice') or '—'}"
            )
        except Exception as exc:
            log_fn(
                "[ERROR] Excel write failed: "
                f"{type(exc).__name__}: {exc}. "
                "Shipment was not written."
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


# Existing app.py imports processor.process_label.
process_label = process_file