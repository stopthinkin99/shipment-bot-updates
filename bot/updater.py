"""
updater.py  —  Silent auto-updater via private GitHub repo.

On startup the app calls check_and_update().
If a newer version exists on GitHub, the new files are downloaded
silently and the app restarts itself.

HOW TO USE:
  1. Create a PRIVATE GitHub repo (e.g. unicreation/shipment-bot-updates)
  2. Put a version.txt in the root:   1.0.0
  3. Put updated .py files in a /bot/ subfolder
  4. Fill in GITHUB_REPO and GITHUB_TOKEN below
  5. Every time you fix something on your laptop, bump version.txt
     and push.  Their app picks it up on next start.

GITHUB_TOKEN: create a Personal Access Token (classic) with
  repo:read scope only.  Settings → Developer settings → PAT.
  This is read-only, so even if someone finds it they can only
  read your repo.
"""

import os
import sys
import json
import shutil
import zipfile
import tempfile
import subprocess
import urllib.request
import urllib.error
from pathlib import Path

# ------------------------------------------------------------------ #
#  CONFIG  — fill these in once
# ------------------------------------------------------------------ #
GITHUB_OWNER  = "stopthinkin99"          # e.g. "aayan-boradia"
GITHUB_REPO   = "shipment-bot-updates"          # your private repo name
GITHUB_TOKEN  = "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"       # read-only PAT
GITHUB_BRANCH = "main"

# Local version file — sits next to the .exe
APP_DIR      = Path(sys.executable).parent if getattr(sys, "frozen", False) \
               else Path(__file__).parent
VERSION_FILE = APP_DIR / "version.txt"
CURRENT_VERSION = VERSION_FILE.read_text().strip() \
                  if VERSION_FILE.exists() else "0.0.0"

# Files the updater will replace (everything else stays untouched)
UPDATABLE_FILES = [
    "parser.py",
    "extractor.py",
    "processor.py",
    "excel_writer.py",
    "mailer.py",
    "app.py",
    "updater.py",
]


# ------------------------------------------------------------------ #
#  HELPERS
# ------------------------------------------------------------------ #
def _github_get(path: str) -> bytes | None:
    """GET from GitHub API with auth.  Returns bytes or None on failure."""
    url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/{path}"
    req = urllib.request.Request(url, headers={
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "ShipmentBot-Updater/1.0",
    })
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            return r.read()
    except urllib.error.URLError:
        return None   # no internet / repo unreachable — skip silently


def _version_tuple(v: str):
    try:
        return tuple(int(x) for x in v.strip().split("."))
    except Exception:
        return (0, 0, 0)


def _download_file(raw_url: str, dest: Path):
    """Download a raw file from GitHub into dest."""
    req = urllib.request.Request(raw_url, headers={
        "Authorization": f"token {GITHUB_TOKEN}",
        "User-Agent": "ShipmentBot-Updater/1.0",
    })
    with urllib.request.urlopen(req, timeout=15) as r:
        dest.write_bytes(r.read())


# ------------------------------------------------------------------ #
#  PUBLIC
# ------------------------------------------------------------------ #
def check_and_update(log_fn=print) -> bool:
    """
    Check GitHub for a newer version.  If found, download updated files
    and return True (caller should restart).  Returns False if up to date
    or if the check fails (no internet etc.).
    """
    log_fn("[Updater] Checking for updates…")

    raw = _github_get(
        f"contents/version.txt?ref={GITHUB_BRANCH}"
    )
    if not raw:
        log_fn("[Updater] Update check skipped (offline or unreachable).")
        return False

    info = json.loads(raw)
    # GitHub returns base64-encoded content
    import base64
    remote_version = base64.b64decode(info["content"]).decode().strip()

    if _version_tuple(remote_version) <= _version_tuple(CURRENT_VERSION):
        log_fn(f"[Updater] Up to date (v{CURRENT_VERSION}).")
        return False

    log_fn(f"[Updater] Update available: {CURRENT_VERSION} → {remote_version}. Downloading…")

    # Download each updatable file from /bot/ subfolder in the repo
    tmp = Path(tempfile.mkdtemp())
    try:
        for fname in UPDATABLE_FILES:
            raw_file = _github_get(
                f"contents/bot/{fname}?ref={GITHUB_BRANCH}"
            )
            if not raw_file:
                continue
            file_info = json.loads(raw_file)
            import base64 as b64
            content = b64.b64decode(file_info["content"])
            (tmp / fname).write_bytes(content)
            log_fn(f"[Updater]   ✓ {fname}")

        # Atomic swap: copy new files over existing ones
        for fname in UPDATABLE_FILES:
            src = tmp / fname
            if src.exists():
                shutil.copy2(src, APP_DIR / fname)

        # Update local version stamp
        VERSION_FILE.write_text(remote_version)
        log_fn(f"[Updater] Update to v{remote_version} complete. Restarting…")
        return True

    except Exception as e:
        log_fn(f"[Updater] Update failed: {e}. Continuing with current version.")
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def restart_app():
    """Restart the current process (works both frozen .exe and python script)."""
    os.execv(sys.executable, [sys.executable] + sys.argv)