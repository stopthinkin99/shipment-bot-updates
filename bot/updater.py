"""
Secure GitHub updater for the Uni Creation Shipment Bot.

Normal startup:
    from updater import sync_all_files
    sync_all_files(log_fn)

One-time token setup on each Windows account/computer:
    py updater.py --set-token

Other maintenance commands:
    py updater.py --test-token
    py updater.py --delete-token

The GitHub token is stored in Windows Credential Manager through keyring.
It is never stored in this source file.
"""

from __future__ import annotations

import argparse
import base64
import getpass
import json
import os
import ssl
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Callable

try:
    import keyring
    from keyring.errors import KeyringError
except ImportError:
    keyring = None

    class KeyringError(Exception):
        pass


GITHUB_OWNER = "stopthinkin99"
GITHUB_REPO = "shipment-bot-updates"
GITHUB_BRANCH = "main"

CREDENTIAL_SERVICE = "UniCreationShipmentBot"
CREDENTIAL_USERNAME = "github-updater-token"
TOKEN_ENVIRONMENT_VARIABLE = "SHIPMENT_BOT_GITHUB_TOKEN"

APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)

SYNCED_FILES = [
    "runtime_paths.py",
    "parser.py",
    "extractor.py",
    "extractor_core.py",
    "po_tracking.py",
    "zales_extractor.py",
    "malka_brinx.py",
    "processor.py",
    "excel_writer.py",
    "mailer.py",
    "cleanup.py",
    "daily_digest.py",
    "email_sender.py",
    "manual_review_dialog.py",
    "fedex_credentials.py",
    "fedex_status_updater.py",
    "fedex_tracking.py",
]


def _create_ssl_context() -> ssl.SSLContext:
    """Create a verified TLS context; never disable certificate checking."""
    try:
        import truststore

        truststore.inject_into_ssl()
        return ssl.create_default_context()
    except Exception:
        pass

    try:
        import certifi

        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


_SSL_CONTEXT = _create_ssl_context()


def _clean_token(token: str | None) -> str:
    value = str(token or "").strip()
    lowered = value.lower()

    if lowered.startswith("bearer "):
        value = value[7:].strip()
    elif lowered.startswith("token "):
        value = value[6:].strip()

    return value


def get_github_token() -> str:
    """Read token from Windows Credential Manager, then env-var fallback."""
    if keyring is not None:
        try:
            stored = keyring.get_password(
                CREDENTIAL_SERVICE,
                CREDENTIAL_USERNAME,
            )
            stored = _clean_token(stored)
            if stored:
                return stored
        except (KeyringError, Exception):
            pass

    return _clean_token(os.environ.get(TOKEN_ENVIRONMENT_VARIABLE, ""))


def save_github_token(token: str) -> None:
    """Save token for the current Windows user in Credential Manager."""
    cleaned = _clean_token(token)

    if not cleaned:
        raise ValueError("The GitHub token cannot be blank.")

    if keyring is None:
        raise RuntimeError(
            "The 'keyring' package is not installed. "
            "Install or bundle keyring before saving the token."
        )

    keyring.set_password(
        CREDENTIAL_SERVICE,
        CREDENTIAL_USERNAME,
        cleaned,
    )


def delete_github_token() -> bool:
    if keyring is None:
        return False

    try:
        keyring.delete_password(
            CREDENTIAL_SERVICE,
            CREDENTIAL_USERNAME,
        )
        return True
    except Exception:
        return False


def _github_headers(token: str) -> dict[str, str]:
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "UniCreation-ShipmentBot-Updater/2.0",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    if token:
        headers["Authorization"] = f"Bearer {token}"

    return headers


def _github_contents_url(path: str) -> str:
    safe_path = "/".join(
        urllib.parse.quote(part, safe="")
        for part in path.replace("\\", "/").split("/")
    )

    return (
        f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
        f"/contents/{safe_path}?ref={urllib.parse.quote(GITHUB_BRANCH, safe='')}"
    )


def _read_error_message(exc: urllib.error.HTTPError) -> str:
    try:
        payload = json.loads(exc.read().decode("utf-8", errors="replace"))
        message = str(payload.get("message", "")).strip()
        if message:
            return message
    except Exception:
        pass

    return str(exc.reason or exc)


def _github_get_file(
    path: str,
    token: str,
    log_fn: Callable[[str], None] = print,
) -> bytes | None:
    if not token:
        log_fn(
            "[Sync] GitHub token is missing. Run updater.py --set-token "
            "under the Windows account that runs Shipment Bot."
        )
        return None

    request = urllib.request.Request(
        _github_contents_url(path),
        headers=_github_headers(token),
        method="GET",
    )

    try:
        with urllib.request.urlopen(
            request,
            timeout=20,
            context=_SSL_CONTEXT,
        ) as response:
            payload = json.loads(response.read().decode("utf-8"))

        encoded = str(payload.get("content", "")).replace("\n", "")
        if not encoded:
            log_fn(f"[Sync] Empty GitHub response for {path}.")
            return None

        return base64.b64decode(encoded, validate=True)

    except urllib.error.HTTPError as exc:
        detail = _read_error_message(exc)

        if exc.code == 401:
            log_fn(
                f"[Sync] Authentication failed for {path}: token invalid, "
                "expired, or revoked."
            )
        elif exc.code == 403:
            log_fn(
                f"[Sync] Access denied for {path}: verify repository access "
                "and Contents: Read permission."
            )
        elif exc.code == 404:
            log_fn(
                f"[Sync] File or repository not found: {path}. "
                "Verify repository access and the bot/ path."
            )
        else:
            log_fn(f"[Sync] GitHub HTTP {exc.code} for {path}: {detail}")

        return None

    except urllib.error.URLError as exc:
        log_fn(f"[Sync] Network/TLS error for {path}: {exc.reason}")
        return None

    except Exception as exc:
        log_fn(
            f"[Sync] Failed to fetch {path}: "
            f"{type(exc).__name__}: {exc}"
        )
        return None


def test_github_access(
    log_fn: Callable[[str], None] = print,
) -> bool:
    token = get_github_token()

    if not token:
        log_fn("[Sync Test] No token is stored for this Windows user.")
        return False

    content = _github_get_file(
        "bot/runtime_paths.py",
        token,
        log_fn,
    )

    if content is None:
        log_fn("[Sync Test] GitHub access failed.")
        return False

    log_fn(
        "[Sync Test] GitHub access is working. "
        "The private repository can be read."
    )
    return True


def _atomic_write(destination: Path, content: bytes) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)

    fd, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.name}.",
        suffix=".tmp",
        dir=str(destination.parent),
    )
    temporary_path = Path(temporary_name)

    try:
        with os.fdopen(fd, "wb") as temporary_file:
            temporary_file.write(content)
            temporary_file.flush()
            os.fsync(temporary_file.fileno())

        os.replace(temporary_path, destination)
    except Exception:
        try:
            temporary_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise


def sync_all_files(
    log_fn: Callable[[str], None] = print,
) -> bool:
    """
    Download the latest modules. Return True only when every file succeeds.
    Existing local copies remain for any file that cannot be downloaded.
    """
    log_fn("[Sync] Pulling latest code from GitHub...")

    token = get_github_token()

    if not token:
        log_fn(
            "[Sync] No GitHub token is stored. "
            "Running with existing local files."
        )
        return False

    successes = 0
    failures = 0

    for filename in SYNCED_FILES:
        content = _github_get_file(
            f"bot/{filename}",
            token,
            log_fn,
        )

        if content is None:
            log_fn(
                f"[Sync]   X {filename} "
                "(not updated; keeping existing copy)"
            )
            failures += 1
            continue

        try:
            _atomic_write(APP_DIR / filename, content)
            log_fn(f"[Sync]   OK {filename}")
            successes += 1
        except Exception as exc:
            log_fn(
                f"[Sync]   X {filename} write failed: "
                f"{type(exc).__name__}: {exc}"
            )
            failures += 1

    if failures == 0 and successes == len(SYNCED_FILES):
        log_fn(f"[Sync] Complete: {successes} files updated successfully.")
        return True

    if successes:
        log_fn(
            f"[Sync] Partial update: {successes} succeeded, "
            f"{failures} failed. Existing copies were kept where needed."
        )
    else:
        log_fn(
            "[Sync] Update failed or unavailable. "
            "Running with existing local files."
        )

    return False


def _set_token_interactively() -> int:
    if keyring is None:
        print("ERROR: keyring is not installed in this Python environment.")
        return 1

    print(
        "The token will be stored in Windows Credential Manager for "
        "the current Windows user."
    )
    token = getpass.getpass("Paste the NEW GitHub token: ").strip()

    if not token:
        print("No token was entered.")
        return 1

    try:
        save_github_token(token)
    except Exception as exc:
        print(f"Could not save the token: {type(exc).__name__}: {exc}")
        return 1

    print("Token saved securely. Testing repository access...")
    return 0 if test_github_access(print) else 1


def _main() -> int:
    parser = argparse.ArgumentParser(
        description="Shipment Bot secure GitHub updater setup."
    )
    parser.add_argument(
        "--set-token",
        action="store_true",
        help="Save a GitHub token in Windows Credential Manager.",
    )
    parser.add_argument(
        "--test-token",
        action="store_true",
        help="Test the stored token against the private repository.",
    )
    parser.add_argument(
        "--delete-token",
        action="store_true",
        help="Delete the stored GitHub token.",
    )

    args = parser.parse_args()
    selected = sum(
        bool(value)
        for value in (
            args.set_token,
            args.test_token,
            args.delete_token,
        )
    )

    if selected != 1:
        parser.print_help()
        return 1

    if args.set_token:
        return _set_token_interactively()

    if args.test_token:
        return 0 if test_github_access(print) else 1

    if args.delete_token:
        if delete_github_token():
            print("Stored GitHub token deleted.")
            return 0

        print("No stored token was deleted.")
        return 1

    return 1


if __name__ == "__main__":
    raise SystemExit(_main())