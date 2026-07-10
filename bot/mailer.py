"""
mailer.py  —  Sends alert emails via local Outlook (win32com).
Zero external SMTP, zero API.  Uses the Outlook desktop app already
installed on the PC — same security boundary as the user's email.
"""

import datetime

# Alert recipient — update this when you have the address
ALERT_TO = "finance@unicreation.com"   # <-- change this

def send_alert(pdf_path: str, data: dict, reason: str):
    """
    Fire an Outlook email telling the user a label needs manual entry.
    Includes everything the bot DID extract so they only have to fill
    in the blanks.
    """
    try:
        import win32com.client as win32
        outlook = win32.Dispatch("Outlook.Application")
        mail    = outlook.CreateItem(0)   # 0 = olMailItem

        mail.To      = ALERT_TO
        mail.Subject = f"[Shipment Bot] Manual Review Needed — {_basename(pdf_path)}"
        mail.Body    = _body(pdf_path, data, reason)
        mail.Send()
        print(f"[MAIL] Alert sent to {ALERT_TO}")

    except Exception as e:
        print(f"[MAIL] Could not send alert email: {e}")
        print(f"[MAIL] Manual review needed for: {pdf_path}")
        print(f"[MAIL] Reason: {reason}")


def _basename(path):
    import os; return os.path.basename(path)


def _body(pdf_path, data, reason):
    extracted = []
    field_map = [
        ("Date",            data.get("date")),
        ("Ship To",         data.get("ship_to")),
        ("Invoice / PO",    data.get("invoice")),
        ("Carrier",         data.get("carrier")),
        ("Tracking Number", data.get("tracking_number")),
        ("Sheet",           data.get("sheet")),
        ("Remark (col G)",  data.get("remark")),
    ]
    for label, value in field_map:
        status = value if value else "⚠ NOT FOUND"
        extracted.append(f"  {label:<20}: {status}")

    lines = [
        "The shipment automation bot could not fully process the label below.",
        "Please open the tracking sheet and enter the missing fields manually.",
        "",
        f"Label file : {pdf_path}",
        f"Reason     : {reason}",
        f"Time       : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "Fields extracted (✓ = confirmed, ⚠ = needs manual entry):",
        *extracted,
        "",
        "Instructions:",
        "  1. Open TRACKING_SHIPMENT.xlsx and go to the correct sheet (see 'Sheet' above).",
        "  2. Add a new row after today's last entry.",
        "  3. Fill in any ⚠ fields by checking the physical label or Diaspark.",
        "  4. If the date is missing, use today's ship date from the label.",
        "",
        "This is an automated message from the Uni Creation Shipment Bot.",
    ]
    return "\n".join(lines)