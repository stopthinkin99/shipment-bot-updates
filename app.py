"""
app.py  —  Uni Creation Shipment Bot
Single-window desktop application. Tkinter only (built into Python).
Auto-starts on Windows login. Always pulls latest bot code from
GitHub on startup — no manual copying, no version comparison.
"""

import sys
import os
from pathlib import Path
import threading
from daily_digest import DigestScheduler, run_daily_digest, load_digest_time, save_digest_time

_APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) \
           else Path(__file__).parent
if str(_APP_DIR) not in sys.path:
    sys.path.insert(0, str(_APP_DIR))

# ------------------------------------------------------------------ #
#  SYNC FIRST — before importing anything that depends on bot files
# ------------------------------------------------------------------ #
from updater import sync_all_files
sync_all_files(print)   # console log; GUI log wired in after window opens

import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import time
import datetime
import winreg
from cleanup import start_cleanup_thread

# ------------------------------------------------------------------ #
#  PATHS
# ------------------------------------------------------------------ #
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) \
          else Path(__file__).parent
CONFIG_FILE  = APP_DIR / "config.json"
VERSION_FILE = APP_DIR / "version.txt"
APP_NAME     = "UniCreationShipmentBot"
# Points to the .exe when frozen, or python app.py when not
APP_EXE      = sys.executable if getattr(sys, "frozen", False) \
               else f'"{sys.executable}" "{__file__}"'


# ------------------------------------------------------------------ #
#  CONFIG
# ------------------------------------------------------------------ #
def load_config() -> dict:
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE) as f:
            return json.load(f)
    return {"watch_folder": "", "excel_path": "", "alert_email": ""}


def save_config(cfg: dict):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)


# ------------------------------------------------------------------ #
#  WINDOWS AUTO-START (registry, no admin needed)
# ------------------------------------------------------------------ #
def _set_autostart(enable: bool):
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Run",
        0, winreg.KEY_SET_VALUE
    )
    if enable:
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, APP_EXE)
    else:
        try:
            winreg.DeleteValue(key, APP_NAME)
        except FileNotFoundError:
            pass
    winreg.CloseKey(key)


def _get_autostart() -> bool:
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_READ
        )
        winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        return False


# ------------------------------------------------------------------ #
#  WATCHER
# ------------------------------------------------------------------ #
class _LabelHandler:
    """Thin wrapper — import processor lazily so updates take effect."""
    def __init__(self, log_fn, excel_path: str):
        self.log = log_fn
        self.excel_path = excel_path

    def dispatch(self, path: str):
        if not path.lower().endswith(".pdf"):
            return
        self.log(f"New label: {os.path.basename(path)}")
        time.sleep(0.8)
        try:
            import importlib, processor
            importlib.reload(processor)
            records = processor.process_label(
                path,
                excel_path=self.excel_path,
                log_fn=self.log,
            )

            if not records:
                self.log(
                    f"✗  {os.path.basename(path)}: processing failed "
                    f"or no records were extracted"
                )
                return

            for i, data in enumerate(records):
                if len(records) > 1:
                    self.log(f"  ── Page {i+1} ──")
                self._log_data(data)

            self.log(f"✓  {os.path.basename(path)}")
        except Exception as e:
            self.log(f"✗  {os.path.basename(path)}: {e}")


    def _log_data(self, data: dict):
        if not data:
            return
        fields = [
            ("Date",     data.get("date")),
            ("Ship To",  data.get("ship_to")),
            ("Invoice",  data.get("invoice")),
            ("Carrier",  data.get("carrier")),
            ("Tracking", data.get("tracking_number")),
            ("Sheet",    data.get("sheet")),
            ("Remark",   data.get("remark")),
        ]
        for label, value in fields:
            shown = value if value else "—"
            self.log(f"    {label:<9}: {shown}")


class WatcherThread(threading.Thread):
    def __init__(self, folder: str, excel_path: str, log_fn):
        super().__init__(daemon=True)
        self.folder = folder
        self.excel_path = excel_path
        self._handler = _LabelHandler(log_fn, excel_path)
        self._stop_evt = threading.Event()
        self._observer = None

    def run(self):
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler

        class _FSHandler(FileSystemEventHandler):
            def __init__(self, dispatch_fn):
                self._dispatch = dispatch_fn
            def on_created(self, event):
                if not event.is_directory:
                    self._dispatch(event.src_path)

        self._observer = Observer()
        self._observer.schedule(
            _FSHandler(self._handler.dispatch),
            self.folder, recursive=False
        )
        self._observer.start()
        while not self._stop_evt.is_set():
            time.sleep(1)
        self._observer.stop()
        self._observer.join()

    def stop(self):
        self._stop_evt.set()


# ------------------------------------------------------------------ #
#  GUI
# ------------------------------------------------------------------ #
class App(tk.Tk):

    # Colour palette
    C_BG      = "#f4f6f9"
    C_HEADER  = "#0d2137"
    C_GREEN   = "#1a7a3c"
    C_RED     = "#a02020"
    C_ACCENT  = "#1558b0"

    def __init__(self):
        super().__init__()
        self.title("Uni Creation — Shipment Bot")
        self.configure(bg=self.C_BG)
        self.resizable(False, False)

        self._cfg     = load_config()
        self._watcher = None
        self._version = VERSION_FILE.read_text().strip() \
                        if VERSION_FILE.exists() else "1.0.0"

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Check for updates then auto-start if already configured
        self.after(500, self._startup)

    # ---------------------------------------------------------------- #
    #  UI BUILD
    # ---------------------------------------------------------------- #
    def _build_ui(self):
        # ── header ──────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=self.C_HEADER)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  Uni Creation  |  Shipment Label Bot",
                 bg=self.C_HEADER, fg="white",
                 font=("Segoe UI", 13, "bold")).pack(side="left", pady=12)
        tk.Label(hdr, text=f"v{self._version}  ",
                 bg=self.C_HEADER, fg="#7090b0",
                 font=("Segoe UI", 8)).pack(side="right", pady=12)

        # ── config ──────────────────────────────────────────────────
        frm = tk.LabelFrame(self, text="Setup",
                            bg=self.C_BG, font=("Segoe UI", 9, "bold"),
                            padx=12, pady=10)
        frm.pack(fill="x", padx=16, pady=(12, 6))

        rows = [
            ("Labels folder",  "var_folder", self._pick_folder),
            ("Excel file",     "var_excel",  self._pick_excel),
        ]
        for i, (label, attr, cmd) in enumerate(rows):
            tk.Label(frm, text=label + ":", bg=self.C_BG,
                     font=("Segoe UI", 9), width=14,
                     anchor="w").grid(row=i, column=0, sticky="w",
                                      pady=(0 if i else 0, 5))
            var = tk.StringVar(value=self._cfg.get(
                attr.replace("var_", "").replace("folder","watch_folder")
                    .replace("excel","excel_path")
                    .replace("email","alert_email"), ""))
            setattr(self, attr, var)
            tk.Entry(frm, textvariable=var, width=44,
                     font=("Segoe UI", 9)).grid(row=i, column=1,
                                                padx=6, pady=(0,5))
            if cmd:
                tk.Button(frm, text="Browse…", command=cmd,
                          font=("Segoe UI", 8),
                          relief="flat", bg="#dde3ec",
                          cursor="hand2").grid(row=i, column=2, pady=(0,5))

        # ── auto-start checkbox ─────────────────────────────────────
        self.var_autostart = tk.BooleanVar(value=_get_autostart())
        tk.Checkbutton(frm, text="Start automatically when Windows starts",
                       variable=self.var_autostart,
                       command=self._toggle_autostart,
                       bg=self.C_BG, font=("Segoe UI", 9),
                       activebackground=self.C_BG).grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(6,0))

        # ── daily summary e-mail ────────────────────────────────────
        # Own frame (packs its children) so it never clashes with the
        # grid()-based Setup frame above.
        dig_frm = tk.Frame(self, bg=self.C_BG)
        dig_frm.pack(fill="x", padx=16, pady=(0, 6))

        # Time entry that remembers itself between launches
        # (saved to digest_settings.json by daily_digest, not config.json).
        self.time_var = tk.StringVar(value=load_digest_time())
        self.time_var.trace_add(
            "write", lambda *a: save_digest_time(self.time_var.get()))
        tk.Label(dig_frm, text="Daily summary time (HH:MM):",
                 bg=self.C_BG, font=("Segoe UI", 9)).pack(side="left")
        tk.Entry(dig_frm, textvariable=self.time_var, width=6,
                 font=("Segoe UI", 9)).pack(side="left", padx=(6, 12))

        tk.Button(dig_frm, text="Send summary now",
                  font=("Segoe UI", 8), relief="flat",
                  bg="#dde3ec", cursor="hand2",
                  command=lambda: threading.Thread(
                      target=lambda: run_daily_digest(
                          self.var_excel.get().strip(), log=self._log),
                      daemon=True).start()
                  ).pack(side="left")

        # Scheduler: fires run_daily_digest once a day at the chosen time.
        self.digest_sched = DigestScheduler(
            get_excel_path=lambda: self.var_excel.get().strip(),
            get_target=lambda: self.time_var.get(),
            log=self._log,
        )
        self.digest_sched.start()

        # ── buttons ─────────────────────────────────────────────────
        btn_frm = tk.Frame(self, bg=self.C_BG)
        btn_frm.pack(fill="x", padx=16, pady=(4, 8))

        self.btn_start = tk.Button(btn_frm, text="▶  Start Watching",
                                   command=self._start,
                                   bg=self.C_GREEN, fg="white", width=18,
                                   font=("Segoe UI", 9, "bold"),
                                   relief="flat", cursor="hand2")
        self.btn_start.pack(side="left", padx=(0, 6))

        self.btn_stop = tk.Button(btn_frm, text="■  Stop",
                                  command=self._stop,
                                  bg=self.C_RED, fg="white", width=10,
                                  font=("Segoe UI", 9, "bold"),
                                  state="disabled",
                                  relief="flat", cursor="hand2")
        self.btn_stop.pack(side="left")

        self.lbl_status = tk.Label(btn_frm, text="⬤  Idle",
                                   fg="#999", bg=self.C_BG,
                                   font=("Segoe UI", 9))
        self.lbl_status.pack(side="right")

        # ── log ─────────────────────────────────────────────────────
        log_frm = tk.LabelFrame(self, text="Activity",
                                bg=self.C_BG, font=("Segoe UI", 9, "bold"),
                                padx=10, pady=6)
        log_frm.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.log_box = tk.Text(log_frm, height=13, state="disabled",
                               bg="#161b22", fg="#c9d1d9",
                               font=("Consolas", 9), relief="flat",
                               wrap="word", insertbackground="white")
        sb = tk.Scrollbar(log_frm, command=self.log_box.yview)
        self.log_box["yscrollcommand"] = sb.set
        self.log_box.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")



    # ---------------------------------------------------------------- #
    #  STARTUP (sync already ran before window opened — just auto-start)
    # ---------------------------------------------------------------- #
    def _startup(self):
        self._log("Synced with GitHub on launch.")
        if self._cfg.get("watch_folder") and self._cfg.get("excel_path"):
            self._start()

    # ---------------------------------------------------------------- #
    #  ACTIONS
    # ---------------------------------------------------------------- #
    def _pick_folder(self):
        d = filedialog.askdirectory(title="Select labels folder")
        if d:
            self.var_folder.set(d)

    def _pick_excel(self):
        f = filedialog.askopenfilename(
            title="Select Excel tracking file",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("All", "*.*")]
        )
        if f:
            self.var_excel.set(f)

    def _toggle_autostart(self):
        _set_autostart(self.var_autostart.get())

    def _save_cfg(self):
        self._cfg = {
            "watch_folder": self.var_folder.get().strip(),
            "excel_path":   self.var_excel.get().strip(),
        }
        save_config(self._cfg)

    def _start(self):
        folder = self.var_folder.get().strip()
        excel  = self.var_excel.get().strip()

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Please choose a valid labels folder.")
            return
        if not excel or not os.path.isfile(excel):
            messagebox.showerror("Error", "Please choose the Excel tracking file.")
            return

        self._save_cfg()
        self._stop()

        self._watcher = WatcherThread(folder, excel, self._log)
        self._watcher.start()

        # Start 30-day cleanup thread (runs now, then every 24 hours)
        start_cleanup_thread(lambda: self.var_folder.get().strip(), self._log)

        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.lbl_status.config(text="⬤  Running", fg=self.C_GREEN)
        self._log(f"Watching: {folder}")

    def _stop(self):
        if self._watcher:
            self._watcher.stop()
            self._watcher = None
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.lbl_status.config(text="⬤  Idle", fg="#999")

    def _on_close(self):
        self._stop()
        if getattr(self, "digest_sched", None):
            self.digest_sched.stop()
        self.destroy()

    # ---------------------------------------------------------------- #
    #  LOG (thread-safe)
    # ---------------------------------------------------------------- #
    def _log(self, msg: str):
        ts   = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}]  {msg}\n"
        def _insert():
            self.log_box.config(state="normal")
            self.log_box.insert("end", line)
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(0, _insert)
        print(line, end="")


if __name__ == "__main__":
    App().mainloop()