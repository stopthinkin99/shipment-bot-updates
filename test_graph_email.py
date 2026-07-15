from email_sender import send_email


send_email(
    recipients=[
        "shipping@unidesignusa.com",
    ],
    subject="Shipment Bot Microsoft Graph Test",
    body="Microsoft Graph delegated email is working.",
)