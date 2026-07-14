"""
updater.py  â€”  Always pulls the LATEST bot files from GitHub on startup.

No version comparison needed. Every time the app opens, it re-downloads
parser.py, extractor.py, processor.py, excel_writer.py, mailer.py,
cleanup.py fresh from the private GitHub repo and overwrites the local
copies. This guarantees the app always runs the newest code you've
pushed, with zero manual copying to the target PC ever required.
"""

import os
import sys
import json
import base64
import ssl
from pathlib import Path
import urllib.request
import urllib.error

# Fix Windows SSL certificate verification issues by using certifi's
# trusted CA bundle instead of relying on the OS store, which can be
# incomplete or misconfigured on some Windows installs.
try:
    import certifi
    _SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())
except ImportError:
    _SSL_CONTEXT = ssl.create_default_context()

# ------------------------------------------------------------------ #
#  CONFIG
# ------------------------------------------------------------------ #
GITHUB_OWNER  = "stopthinkin99"
GITHUB_REPO   = "shipment-bot-updates"
GITHUB_TOKEN  = "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"
GITHUB_BRANCH = "main"

APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) \
          else Path(__file__).parent

# Every file that should always be pulled fresh from GitHub on startup.
# NOTE: updater.py itself is intentionally NOT in this list â€” the app
# can't safely overwrite the file it's currently executing from disk
# mid-run on Windows. Update updater.py manually if it ever changes.
SYNCED_FILES = [
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
]


# ------------------------------------------------------------------ #
#  GITHUB HELPERS
# ------------------------------------------------------------------ #
def _github_get_file(path: str):
    """Fetch a single file's raw bytes from the repo. None on failure."""
    url = (f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
           f"/contents/{path}?ref={GITHUB_BRANCH}")
    req = urllib.request.Request(url, headers={
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "ShipmentBot-Updater/1.0",
    })
    try:
        with urllib.request.urlopen(req, timeout=10, context=_SSL_CONTEXT) as r:
            info = json.loads(r.read())
            return base64.b64decode(info["content"])
    except Exception:
        return None


# ------------------------------------------------------------------ #
#  PUBLIC
# ------------------------------------------------------------------ #
def sync_all_files(log_fn=print) -> bool:
    """
    Download the latest version of every file in SYNCED_FILES from
    GitHub and overwrite the local copy. Returns True if at least one
    file was successfully synced, False if GitHub was unreachable
    (in which case the app just runs whatever is already on disk).
    """
    log_fn("[Sync] Pulling latest code from GitHubâ€¦")

    any_success = False
    for fname in SYNCED_FILES:
        content = _github_get_file(f"bot/{fname}")
        if content is None:
            log_fn(f"[Sync]   âœ— {fname} (unreachable â€” keeping existing copy)")
            continue
        try:
            (APP_DIR / fname).write_bytes(content)
            log_fn(f"[Sync]   âœ“ {fname}")
            any_success = True
        except Exception as e:
            log_fn(f"[Sync]   âœ— {fname} write failed: {e}")

    if any_success:
        log_fn("[Sync] Up to date with GitHub.")
    else:
        log_fn("[Sync] Offline â€” running with existing local files.")

    return any_success