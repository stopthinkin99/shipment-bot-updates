"""
mailer.py
---------
Sends manual-review alerts through Microsoft Graph delegated authentication.
Outlook desktop is not required.

This module must be called by processor.py when required extraction fields
are missing or the target Excel sheet cannot be determined.
"""

import datetime
import os

from email_sender import send_email


# ------------------------------------------------------------------ #
#  RECIPIENT CONFIGURATION
# ------------------------------------------------------------------ #
MANUAL_REVIEW_RECIPIENTS = [
    "shipping@unidesignusa.com",
]


def send_alert(
    pdf_path: str,
    data: dict,
    reason: str,
    log_fn=print,
) -> bool:
    """
    Send a manual-review notification.

    Returns True when Microsoft Graph accepts the message and False when
    sending fails. The extraction workflow should continue either way.
    """
    try:
        recipients = [
            address.strip()
            for address in MANUAL_REVIEW_RECIPIENTS
            if address and address.strip()
        ]

        if not recipients:
            raise ValueError(
                "No manual-review recipient addresses are configured."
            )

        send_email(
            recipients=recipients,
            subject=(
                "[Shipment Bot] Manual Review — "
                f"{os.path.basename(pdf_path)}"
            ),
            body=_body(pdf_path, data, reason),
            html=False,
            log_fn=log_fn,
        )

        log_fn(
            "[MAIL] Manual-review alert sent to "
            + ", ".join(recipients)
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


def _body(pdf_path, data, reason):
    """Build the plain-text manual-review email body."""
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
        "The shipment bot could not fully process the label below.",
        "Please open the tracking workbook and fill in the missing fields.",
        "",
        f"Label file : {pdf_path}",
        f"Reason     : {reason}",
        f"Time       : {datetime.datetime.now():%Y-%m-%d %H:%M}",
        "",
        "Fields extracted by the bot:",
        *extracted,
        "",
        "Instructions:",
        "  1. Open the selected shipment-tracking workbook.",
        "  2. Go to the correct entity sheet.",
        "  3. Add or correct the shipment row.",
        "  4. Fill every field marked *** NOT FOUND *** from the label.",
        "",
        "This is an automated message from the Uni Creation Shipment Bot.",
    ]

    return "\n".join(lines)