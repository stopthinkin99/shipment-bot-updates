"""
mailer.py
---------
Sends manual-review alerts through Microsoft Graph and attaches the
original shipment-label PDF.
"""

import datetime
import os

from email_sender import send_email

MANUAL_REVIEW_RECIPIENTS = [
    "shipping@unidesignusa.com",
]


def send_alert(
    pdf_path: str,
    data: dict,
    reason: str,
    log_fn=print,
) -> bool:
    try:
        recipients = [
            address.strip()
            for address in MANUAL_REVIEW_RECIPIENTS
            if address and address.strip()
        ]

        if not recipients:
            raise ValueError("No manual-review recipient addresses are configured.")

        if not os.path.isfile(pdf_path):
            raise FileNotFoundError(
                "The label PDF could not be attached because it was not found: "
                f"{pdf_path}"
            )

        send_email(
            recipients=recipients,
            subject=(
                "[Shipment Bot] Manual Review — "
                f"{os.path.basename(pdf_path)}"
            ),
            body=_body(pdf_path, data, reason),
            html=False,
            attachments=[pdf_path],
            log_fn=log_fn,
        )

        log_fn(
            "[MAIL] Manual-review alert sent to "
            + ", ".join(recipients)
            + f" with attachment {os.path.basename(pdf_path)}"
        )
        return True

    except Exception as exc:
        log_fn(
            "[MAIL ERROR] Manual-review alert could not be sent: "
            f"{type(exc).__name__}: {exc}"
        )
        log_fn(f"[MAIL] File requiring review: {pdf_path}")
        log_fn(f"[MAIL] Reason: {reason}")
        return False


def _body(pdf_path: str, data: dict, reason: str) -> str:
    data = data or {}

    field_map = [
        ("Date", data.get("date")),
        ("Ship To", data.get("ship_to")),
        ("Invoice / PO", data.get("invoice")),
        ("Carrier", data.get("carrier")),
        ("Tracking Number", data.get("tracking_number")),
        ("Sheet", data.get("sheet")),
        ("Remark (col G)", data.get("remark")),
    ]

    extracted = []
    for label, value in field_map:
        shown = (
            str(value).strip()
            if value is not None and str(value).strip()
            else "*** NOT FOUND — enter manually ***"
        )
        extracted.append(f"  {label:<20}: {shown}")

    lines = [
        "The shipment bot could not fully process the attached label.",
        "Please review the attached PDF and fill in the missing information manually.",
        "",
        f"Attached label: {os.path.basename(pdf_path)}",
        f"Original path : {pdf_path}",
        f"Reason        : {reason}",
        f"Time          : {datetime.datetime.now():%Y-%m-%d %H:%M}",
        "",
        "Fields extracted by the bot:",
        *extracted,
        "",
        "Instructions:",
        "  1. Open the attached PDF label.",
        "  2. Open the selected shipment-tracking workbook.",
        "  3. Go to the correct entity sheet.",
        "  4. Add or correct the shipment row.",
        "  5. Fill every field marked *** NOT FOUND ***.",
        "",
        "This is an automated message from the Uni Creation Shipment Bot.",
    ]

    return "\n".join(lines)