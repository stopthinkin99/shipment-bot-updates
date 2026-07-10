"""
mailer.py  —  Sends alert emails via local Outlook (win32com).
"""

import datetime

ALERT_TO = "shipping@unidesignusa.com"


def send_alert(pdf_path: str, data: dict, reason: str):
    try:
        import win32com.client as win32

        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To      = ALERT_TO
        mail.Subject = f"[Shipment Bot] Manual Review — {_basename(pdf_path)}"
        mail.Body    = _body(pdf_path, data, reason)

        # SentOnBehalfOfName not needed — just send directly
        mail.Send()

        print(f"[MAIL] Alert sent to {ALERT_TO}")

    except Exception as e:
        # Fallback: save as draft so nothing is lost
        try:
            import win32com.client as win32
            outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To      = ALERT_TO
            mail.Subject = f"[Shipment Bot] Manual Review — {_basename(pdf_path)}"
            mail.Body    = _body(pdf_path, data, reason)
            mail.Save()   # saves to Drafts folder
            print(f"[MAIL] Could not send — saved to Outlook Drafts instead.")
        except Exception as e2:
            print(f"[MAIL] Email failed entirely: {e2}")

        print(f"[MAIL] Manual review needed for: {pdf_path}")
        print(f"[MAIL] Reason: {reason}")


def _basename(path):
    import os
    return os.path.basename(path)


def _body(pdf_path, data, reason):
    field_map = [
        ("Date",            data.get("date")),
        ("Ship To",         data.get("ship_to")),
        ("Invoice / PO",    data.get("invoice")),
        ("Carrier",         data.get("carrier")),
        ("Tracking Number", data.get("tracking_number")),
        ("Sheet",           data.get("sheet")),
        ("Remark (col G)",  data.get("remark")),
    ]

    extracted = []
    for label, value in field_map:
        status = value if value else "*** NOT FOUND — enter manually ***"
        extracted.append(f"  {label:<20}: {status}")

    lines = [
        "The shipment bot could not fully process the label below.",
        "Please open the tracking sheet and fill in the missing fields.",
        "",
        f"Label file : {pdf_path}",
        f"Reason     : {reason}",
        f"Time       : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "Fields extracted by the bot:",
        *extracted,
        "",
        "Instructions:",
        "  1. Open TRACKING_SHIPMENT.xlsx and go to the correct sheet.",
        "  2. Add a new row after today's last entry.",
        "  3. Fill in any *** NOT FOUND *** fields from the physical label.",
        "",
        "This is an automated message from the Uni Creation Shipment Bot.",
    ]
    return "\n".join(lines)