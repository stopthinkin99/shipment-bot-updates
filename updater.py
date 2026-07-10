"""
updater.py  —  Always pulls the LATEST bot files from GitHub on startup.
"""

import os
import sys
import json
import base64
from pathlib import Path
import urllib.request
import urllib.error

GITHUB_OWNER  = "stopthinkin99"
GITHUB_REPO   = "shipment-bot-updates"
GITHUB_TOKEN  = "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"
GITHUB_BRANCH = "main"

APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) \
          else Path(__file__).parent

SYNCED_FILES = [
    "parser.py",
    "extractor.py",
    "processor.py",
    "excel_writer.py",
    "mailer.py",
    "cleanup.py",
]


def _github_get_file(path: str):
    url = (f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
           f"/contents/{path}?ref={GITHUB_BRANCH}")
    req = urllib.request.Request(url, headers={
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "ShipmentBot-Updater/1.0",
    })
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            info = json.loads(r.read())
            return base64.b64decode(info["content"])
    except Exception:
        return None


def sync_all_files(log_fn=print) -> bool:
    log_fn("[Sync] Pulling latest code from GitHub...")

    any_success = False
    for fname in SYNCED_FILES:
        content = _github_get_file(f"bot/{fname}")
        if content is None:
            log_fn(f"[Sync]   FAILED {fname} (unreachable - keeping existing copy)")
            continue
        try:
            (APP_DIR / fname).write_bytes(content)
            log_fn(f"[Sync]   OK {fname}")
            any_success = True
        except Exception as e:
            log_fn(f"[Sync]   FAILED write {fname}: {e}")

    if any_success:
        log_fn("[Sync] Up to date with GitHub.")
    else:
        log_fn("[Sync] Offline - running with existing local files.")

    return any_success
