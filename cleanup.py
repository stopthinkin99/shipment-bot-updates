"""
cleanup.py  —  Auto-deletes label PDFs older than 30 days.

Uses the file's creation time (when it was saved into the watch folder).
Runs once on startup, then every 24 hours in a background thread.
"""

import os
import time
import threading
import datetime

EXPIRY_DAYS   = 30
CHECK_EVERY_S = 60 * 60 * 24   # run once every 24 hours


def _delete_expired(folder: str, log_fn=print):
    """Delete any .pdf in folder whose creation time is > 30 days ago."""
    if not folder or not os.path.isdir(folder):
        return

    cutoff = time.time() - (EXPIRY_DAYS * 86400)
    deleted = 0

    for fname in os.listdir(folder):
        if not fname.lower().endswith(".pdf"):
            continue

        fpath = os.path.join(folder, fname)
        try:
            # st_ctime on Windows = file creation time
            created = os.path.getctime(fpath)
            if created < cutoff:
                os.remove(fpath)
                age = (time.time() - created) / 86400
                log_fn(f"[Cleanup] Deleted {fname} (aged {age:.0f} days)")
                deleted += 1
        except Exception as e:
            log_fn(f"[Cleanup] Could not delete {fname}: {e}")

    if deleted == 0:
        log_fn(f"[Cleanup] No expired labels found.")
    else:
        log_fn(f"[Cleanup] {deleted} expired label(s) deleted.")


def start_cleanup_thread(folder_fn, log_fn=print):
    """
    Start a background thread that:
      - runs immediately on startup
      - then repeats every 24 hours

    folder_fn: a callable that returns the current watch folder path
               (passed as a function so it picks up config changes)
    """
    def _loop():
        while True:
            _delete_expired(folder_fn(), log_fn)
            time.sleep(CHECK_EVERY_S)

    t = threading.Thread(target=_loop, daemon=True)
    t.start()
    return t