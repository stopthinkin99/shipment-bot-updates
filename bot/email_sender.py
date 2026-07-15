import json
import os
import sys
from pathlib import Path
from typing import Iterable

import msal
import requests


# ------------------------------------------------------------------ #
#  MICROSOFT ENTRA CONFIGURATION
# ------------------------------------------------------------------ #

TENANT_ID = "5f22a7de-6c48-4b32-87dd-4f8e0faf6d09"
CLIENT_ID = "5874c908-c525-4779-8c4a-33cf5bdbefed"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Delegated permissions requested for the signed-in user.
SCOPES = [
    "Mail.Send",
    "User.Read",
]


# ------------------------------------------------------------------ #
#  TOKEN CACHE
# ------------------------------------------------------------------ #

def _application_data_directory() -> Path:
    """
    Store the login cache in the user's local application-data folder,
    rather than inside Program Files.
    """
    local_app_data = os.environ.get("LOCALAPPDATA")

    if local_app_data:
        folder = Path(local_app_data) / "UniCreation" / "ShipmentBot"
    else:
        folder = Path.home() / ".un_creation_shipment_bot"

    folder.mkdir(parents=True, exist_ok=True)
    return folder


TOKEN_CACHE_FILE = _application_data_directory() / "msal_token_cache.json"


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()

    if TOKEN_CACHE_FILE.exists():
        try:
            cache.deserialize(
                TOKEN_CACHE_FILE.read_text(encoding="utf-8")
            )
        except Exception:
            # A corrupt cache should not prevent the user from signing in.
            pass

    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(
            cache.serialize(),
            encoding="utf-8",
        )


# ------------------------------------------------------------------ #
#  AUTHENTICATION
# ------------------------------------------------------------------ #

def get_access_token(log_fn=print) -> str:
    cache = _load_cache()

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    # First try the cached account so normal sends happen silently.
    accounts = app.get_accounts()

    result = None
    if accounts:
        result = app.acquire_token_silent(
            scopes=SCOPES,
            account=accounts[0],
        )

    # No usable cached token: ask the user to complete browser sign-in.
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

    access_token = result.get("access_token")
    if access_token:
        return access_token

    error = result.get("error", "unknown_error")
    description = result.get(
        "error_description",
        "Microsoft did not return an access token.",
    )

    raise RuntimeError(f"{error}: {description}")


# ------------------------------------------------------------------ #
#  SEND EMAIL
# ------------------------------------------------------------------ #

def send_email(
    recipients: Iterable[str],
    subject: str,
    body: str,
    *,
    html: bool = False,
    cc: Iterable[str] | None = None,
    bcc: Iterable[str] | None = None,
    log_fn=print,
) -> bool:
    to_list = [x.strip() for x in recipients if x and x.strip()]
    cc_list = [x.strip() for x in (cc or []) if x and x.strip()]
    bcc_list = [x.strip() for x in (bcc or []) if x and x.strip()]

    if not to_list:
        raise ValueError("No email recipients were provided.")

    token = get_access_token(log_fn=log_fn)

    def recipient_objects(addresses: list[str]) -> list[dict]:
        return [
            {
                "emailAddress": {
                    "address": address,
                }
            }
            for address in addresses
        ]

    message = {
        "subject": subject,
        "body": {
            "contentType": "HTML" if html else "Text",
            "content": body,
        },
        "toRecipients": recipient_objects(to_list),
    }

    if cc_list:
        message["ccRecipients"] = recipient_objects(cc_list)

    if bcc_list:
        message["bccRecipients"] = recipient_objects(bcc_list)

    response = requests.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={
            "message": message,
            "saveToSentItems": True,
        },
        timeout=30,
    )

    if response.status_code != 202:
        raise RuntimeError(
            f"Microsoft Graph send failed "
            f"({response.status_code}): {response.text}"
        )

    log_fn(
        "[MAIL] Sent successfully to "
        + ", ".join(to_list)
    )
    return True