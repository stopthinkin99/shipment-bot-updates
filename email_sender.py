"""
email_sender.py
---------------
Sends email through Microsoft Graph using delegated device-code login.
Supports To/CC/BCC, HTML or text bodies, and file attachments.
"""

import base64
import json
import mimetypes
import os
from pathlib import Path
from typing import Iterable

import msal
import requests

TENANT_ID = "5f22a7de-6c48-4b32-87dd-4f8e0faf6d09"
CLIENT_ID = "5874c908-c525-4779-8c4a-33cf5bdbefed"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Send", "User.Read"]


def _application_data_directory() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        folder = Path(local_app_data) / "UniCreation" / "ShipmentBot"
    else:
        folder = Path.home() / ".uni_creation_shipment_bot"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


TOKEN_CACHE_FILE = _application_data_directory() / "msal_token_cache.json"


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        try:
            cache.deserialize(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")


def get_access_token(log_fn=print) -> str:
    if TENANT_ID.startswith("PASTE-") or CLIENT_ID.startswith("PASTE-"):
        raise RuntimeError("TENANT_ID and CLIENT_ID are not configured in email_sender.py.")

    cache = _load_cache()
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes=SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(
                "Microsoft sign-in could not be started: "
                + json.dumps(flow, indent=2)
            )
        log_fn("[MAIL LOGIN] Microsoft sign-in is required.")
        log_fn(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    _save_cache(cache)

    token = result.get("access_token")
    if token:
        return token

    raise RuntimeError(
        f"{result.get('error', 'unknown_error')}: "
        f"{result.get('error_description', 'Microsoft did not return an access token.')}"
    )


def _clean_addresses(addresses: Iterable[str] | None) -> list[str]:
    return [
        str(address).strip()
        for address in (addresses or [])
        if address and str(address).strip()
    ]


def _recipient_objects(addresses: list[str]) -> list[dict]:
    return [
        {"emailAddress": {"address": address}}
        for address in addresses
    ]


def _attachment_objects(attachment_paths: Iterable[str] | None) -> list[dict]:
    attachments = []

    for supplied_path in attachment_paths or []:
        path = Path(supplied_path)
        if not path.is_file():
            raise FileNotFoundError(f"Email attachment was not found: {path}")

        content_type, _ = mimetypes.guess_type(path.name)
        if not content_type:
            content_type = "application/octet-stream"

        attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": path.name,
            "contentType": content_type,
            "contentBytes": base64.b64encode(path.read_bytes()).decode("ascii"),
        })

    return attachments


def send_email(
    recipients: Iterable[str],
    subject: str,
    body: str,
    *,
    html: bool = False,
    cc: Iterable[str] | None = None,
    bcc: Iterable[str] | None = None,
    attachments: Iterable[str] | None = None,
    log_fn=print,
) -> bool:
    to_list = _clean_addresses(recipients)
    cc_list = _clean_addresses(cc)
    bcc_list = _clean_addresses(bcc)

    if not to_list:
        raise ValueError("No email recipients were provided.")

    token = get_access_token(log_fn=log_fn)

    message = {
        "subject": str(subject),
        "body": {
            "contentType": "HTML" if html else "Text",
            "content": str(body),
        },
        "toRecipients": _recipient_objects(to_list),
    }

    if cc_list:
        message["ccRecipients"] = _recipient_objects(cc_list)
    if bcc_list:
        message["bccRecipients"] = _recipient_objects(bcc_list)

    attachment_list = _attachment_objects(attachments)
    if attachment_list:
        message["attachments"] = attachment_list

    response = requests.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"message": message, "saveToSentItems": True},
        timeout=60,
    )

    if response.status_code != 202:
        raise RuntimeError(
            f"Microsoft Graph send failed ({response.status_code}): {response.text}"
        )

    suffix = (
        f" with {len(attachment_list)} attachment(s)"
        if attachment_list else ""
    )
    log_fn("[MAIL] Sent successfully to " + ", ".join(to_list) + suffix)
    return True