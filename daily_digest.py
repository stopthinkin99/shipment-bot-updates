"""
daily_digest.py
---------------
Scheduled daily summary e-mail for the shipping-label bot.

At a user-chosen time each day this module:
  1. opens the same tracking workbook the bot already writes to,
  2. walks each entity sheet (UNI, EMBY, FENIX, SOL),
  3. pulls every row whose date == today,
  4. builds one e-mail per sheet listing those shipments
     (company, invoice, tracking, carrier), and
  5. sends it to the matching Outlook Contact Group.

Column layout is kept in lock-step with excel_writer.py:
  A=Date  B=Ship-to  C=Invoice  D=Carrier  E=Tracking  G=Remark
Reuses the same win32com Outlook path as mailer.py and the same
openpyxl workbook as excel_writer.py, so it drops into the existing app.

app.py wire-up is in the comment block at the bottom of this file.
"""

import os
import json
import shutil
import tempfile
import threading
import time
from datetime import datetime, date

import openpyxl

try:
    import win32com.client  # same dependency mailer.py already uses
except ImportError:  # allows import on a dev machine without Outlook
    win32com = None


# ------------------------------------------------------------------ #
#  CONFIG  --  the only things you may need to touch
# ------------------------------------------------------------------ #

# Sheet name (as it appears in the workbook, case-insensitive)  ->
# Outlook Contact Group display name. Resolved by name at send time.
SHEET_TO_GROUP = {
    "UNI":   "Uni-Shipping",
    "EMBY":  "Emby-Shipping",
    "FENIX": "Fenix-Shipping",
    "SOL":   "Sol-Shipping",
}

# Column positions as 0-based tuple indexes (read side uses iter_rows,
# which is 0-based). Matches the sheet header row:
#   A DATE | B SHIP TO | C INVOICE/MEMO | D CARRIER | E TRACKING NO
# (F REMARK/status, G EXTRA, H DELIVERY DATE are not needed here.)
IDX_DATE     = 0   # A  DATE
IDX_COMPANY  = 1   # B  SHIP TO
IDX_INVOICE  = 2   # C  INVOICE/MEMO
IDX_CARRIER  = 3   # D  CARRIER
IDX_TRACKING = 4   # E  TRACKING NO

# Set True while testing: opens each e-mail in Outlook for you to eyeball
# instead of sending it. Flip to False for unattended sending.
REVIEW_BEFORE_SEND = True

# If a sheet has no shipments today, skip it silently (True) or send a
# short "no shipments today" note anyway (False).
SKIP_EMPTY_SHEETS = True


# ------------------------------------------------------------------ #
#  TIME PERSISTENCE  (remembers the send time between launches,
#  in its own file -- no need to touch your app's config)
# ------------------------------------------------------------------ #

_SETTINGS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "digest_settings.json"
)


def load_digest_time(default="17:30"):
    """Return the saved HH:MM time, or `default` if none saved yet."""
    try:
        with open(_SETTINGS_FILE) as f:
            return json.load(f).get("digest_time", default)
    except Exception:  # noqa -- missing/corrupt file -> use default
        return default


def save_digest_time(value):
    """Persist the HH:MM time so it survives restarts."""
    try:
        with open(_SETTINGS_FILE, "w") as f:
            json.dump({"digest_time": value}, f)
    except Exception:  # noqa -- best effort, never crash the GUI
        pass


# ------------------------------------------------------------------ #
#  WORKBOOK READING
# ------------------------------------------------------------------ #

def _open_workbook(excel_path, retries=4, wait=1.5):
    """
    Load the workbook read-only. If Excel has the file locked, copy it to
    a temp file and read the copy instead (same defensive idea as the
    writer's retry logic). Returns (workbook, temp_path_or_None).
    """
    for _ in range(retries):
        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
            return wb, None
        except (PermissionError, OSError):
            time.sleep(wait)

    # Still locked -> read a copy.
    tmp = os.path.join(tempfile.gettempdir(),
                       f"digest_{os.path.basename(excel_path)}")
    shutil.copy2(excel_path, tmp)
    wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
    return wb, tmp


def _as_date(value):
    """Normalize a cell value to a date, or None."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d%b%y", "%d%b%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _get(row, idx):
    if idx < len(row) and row[idx] is not None:
        return str(row[idx]).strip()
    return ""


def read_todays_shipments(ws, today):
    """
    Return a list of shipment dicts on `ws` dated `today`.

    Header-position-agnostic: the merged title rows, the "DATE" header
    row, blue date-divider rows and blank spacers all fail the "column A
    is a real date == today" test, so they're skipped without any
    hard-coded start row. A row also needs a company or tracking value.
    """
    shipments = []
    for row in ws.iter_rows(values_only=True):
        if row is None or not any(c is not None for c in row):
            continue
        if _as_date(row[IDX_DATE] if len(row) > IDX_DATE else None) != today:
            continue
        company = _get(row, IDX_COMPANY)
        tracking = _get(row, IDX_TRACKING)
        if not company and not tracking:
            continue  # divider / spacer
        shipments.append({
            "company":  company,
            "invoice":  _get(row, IDX_INVOICE),
            "tracking": tracking,
            "carrier":  _get(row, IDX_CARRIER),
        })
    return shipments


def collect_by_sheet(excel_path, today=None):
    """
    Read every configured sheet and return {SHEET_NAME: [shipment, ...]},
    containing only sheets with shipments today (unless SKIP_EMPTY_SHEETS
    is False).
    """
    today = today or date.today()
    wb, tmp = _open_workbook(excel_path)
    result = {}
    try:
        titles = {t.upper(): t for t in wb.sheetnames}
        for sheet_key in SHEET_TO_GROUP:
            actual = titles.get(sheet_key.upper())
            if not actual:
                continue
            shipments = read_todays_shipments(wb[actual], today)
            if shipments or not SKIP_EMPTY_SHEETS:
                result[sheet_key] = shipments
    finally:
        wb.close()
        if tmp and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    return result


# ------------------------------------------------------------------ #
#  E-MAIL BUILDING + SENDING
# ------------------------------------------------------------------ #

def _build_html(sheet_key, shipments, today):
    day = today.strftime("%m/%d/%Y")
    if not shipments:
        return f"<p>No {sheet_key} shipments went out on {day}.</p>"

    head = (
        "<tr style='background:#5B9BD5;color:#fff'>"
        "<th align='left'>Company</th>"
        "<th align='left'>Invoice</th>"
        "<th align='left'>Tracking #</th>"
        "<th align='left'>Carrier</th>"
        "</tr>"
    )
    body_rows = "".join(
        "<tr>"
        f"<td>{s['company']}</td>"
        f"<td>{s['invoice']}</td>"
        f"<td>{s['tracking']}</td>"
        f"<td>{s['carrier']}</td>"
        "</tr>"
        for s in shipments
    )
    return (
        f"<p>{len(shipments)} {sheet_key} shipment(s) went out on {day}:</p>"
        "<table border='1' cellpadding='6' cellspacing='0' "
        "style='border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;"
        "font-size:13px'>"
        f"{head}{body_rows}"
        "</table>"
    )


def _get_outlook():
    if win32com is None:
        raise RuntimeError("win32com not available -- Outlook cannot be reached.")
    return win32com.client.Dispatch("Outlook.Application")


def send_group_digest(sheet_key, shipments, today, outlook=None):
    """Send (or Display) one digest e-mail for a single sheet/group."""
    group = SHEET_TO_GROUP.get(sheet_key)
    if not group:
        return False, f"No Outlook group mapped for {sheet_key}"

    outlook = outlook or _get_outlook()
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.Subject = f"Shipments Sent Today - {sheet_key} - {today.strftime('%m/%d/%Y')}"
    mail.HTMLBody = _build_html(sheet_key, shipments, today)

    recip = mail.Recipients.Add(group)      # Contact Group display name
    recip.Resolve()
    if not mail.Recipients.ResolveAll():
        return False, f"{sheet_key}: could not resolve group '{group}' in Outlook"

    if REVIEW_BEFORE_SEND:
        mail.Display()
        return True, f"{sheet_key}: opened for review ({len(shipments)} rows)"
    mail.Send()
    return True, f"{sheet_key}: sent to {group} ({len(shipments)} rows)"


def run_daily_digest(excel_path, today=None, log=print):
    """
    Full run: read all sheets, send one e-mail per sheet with shipments.
    `log` is any callable(str); pass your GUI logger here. Returns a list
    of status strings.
    """
    today = today or date.today()
    try:
        by_sheet = collect_by_sheet(excel_path, today)
    except Exception as e:  # noqa
        msg = f"Digest read failed: {e}"
        log(msg)
        return [msg]

    if not by_sheet:
        msg = f"No shipments found for {today.strftime('%m/%d/%Y')} on any sheet."
        log(msg)
        return [msg]

    outlook = _get_outlook()
    results = []
    for sheet_key, shipments in by_sheet.items():
        try:
            _, info = send_group_digest(sheet_key, shipments, today, outlook)
        except Exception as e:  # noqa
            info = f"{sheet_key}: FAILED -- {e}"
        log(info)
        results.append(info)
    return results


# ------------------------------------------------------------------ #
#  SCHEDULER
# ------------------------------------------------------------------ #

class DigestScheduler(threading.Thread):
    """
    Background thread that fires run_daily_digest once per day at HH:MM.

    - Editable at runtime: change the value returned by get_target and it
      takes effect on the next poll.
    - Fires at most once per calendar day.
    - Stop cleanly with .stop().
    """

    def __init__(self, get_excel_path, get_target, log=print, poll_seconds=20):
        super().__init__(daemon=True)
        self.get_excel_path = get_excel_path
        self.get_target = get_target      # returns "HH:MM"
        self.log = log
        self.poll_seconds = poll_seconds
        self._stop = threading.Event()
        self._last_fired = None           # date it last ran

    def stop(self):
        self._stop.set()

    def run(self):
        self.log("Daily digest scheduler started.")
        while not self._stop.is_set():
            target = (self.get_target() or "").strip()
            now = datetime.now()
            if target and now.strftime("%H:%M") == target and self._last_fired != now.date():
                self._last_fired = now.date()
                path = self.get_excel_path()
                if not path or not os.path.exists(path):
                    self.log("Digest skipped: Excel path not set / missing.")
                else:
                    self.log(f"Running daily digest for {now.strftime('%m/%d/%Y')} ...")
                    run_daily_digest(path, today=now.date(), log=self.log)
            self._stop.wait(self.poll_seconds)
        self.log("Daily digest scheduler stopped.")


# ------------------------------------------------------------------ #
#  APP.PY WIRE-UP  (paste the relevant bits into app.py)
# ------------------------------------------------------------------ #
#
#  import threading
#  import tkinter as tk
#  from daily_digest import (DigestScheduler, run_daily_digest,
#                            load_digest_time, save_digest_time)
#
#  # --- in your GUI build ---
#  # time entry that remembers itself between launches:
#  time_var = tk.StringVar(value=load_digest_time())          # loads saved time
#  time_var.trace_add("write", lambda *a: save_digest_time(time_var.get()))
#  tk.Label(frame, text="Daily summary time (HH:MM):").pack(side="left")
#  tk.Entry(frame, textvariable=time_var, width=6).pack(side="left")
#
#  # thread-safe logging back into your existing activity log:
#  def gui_log(msg):
#      root.after(0, lambda: activity_log_append(msg))        # your log fn
#
#  # start the scheduler once, when the app starts (or on Start):
#  self.digest_sched = DigestScheduler(
#      get_excel_path=lambda: excel_path_var.get(),           # your existing var
#      get_target=lambda: time_var.get(),
#      log=gui_log,
#  )
#  self.digest_sched.start()
#
#  # optional "Send summary now" button for testing:
#  tk.Button(frame, text="Send summary now",
#            command=lambda: threading.Thread(
#                target=lambda: run_daily_digest(excel_path_var.get(), log=gui_log),
#                daemon=True).start()
#           ).pack(side="left")
#
#  # OPTIONAL clean shutdown (the thread is a daemon, so this isn't required):
#  #   def on_close():
#  #       self.digest_sched.stop()
#  #       root.destroy()
#  #   root.protocol("WM_DELETE_WINDOW", on_close)