"""Secure UNI/FENIX FedEx credentials in Windows Credential Manager."""
from __future__ import annotations
import tkinter as tk
from tkinter import messagebox
import keyring

SERVICE_NAME = "UniCreationShipmentBot.FedEx"
PROFILES = ("UNI", "FENIX")


def _username(profile: str, field: str) -> str:
    profile = profile.strip().upper()
    field = field.strip().upper()
    if profile not in PROFILES:
        raise ValueError(f"Unsupported profile: {profile}")
    if field not in {"API_KEY", "SECRET_KEY"}:
        raise ValueError(field)
    return f"FEDEX_{profile}_{field}"


def save_credentials(profile: str, api_key: str, secret_key: str) -> None:
    profile = profile.strip().upper()
    api_key = api_key.strip()
    secret_key = secret_key.strip()
    if not api_key or not secret_key:
        raise ValueError(f"{profile}: API Key and Secret Key are required.")
    keyring.set_password(SERVICE_NAME, _username(profile, "API_KEY"), api_key)
    keyring.set_password(SERVICE_NAME, _username(profile, "SECRET_KEY"), secret_key)


def load_credentials(profile: str):
    profile = profile.strip().upper()
    api = keyring.get_password(SERVICE_NAME, _username(profile, "API_KEY")) or ""
    secret = keyring.get_password(SERVICE_NAME, _username(profile, "SECRET_KEY")) or ""
    return api.strip(), secret.strip()


def has_credentials(profile: str) -> bool:
    api, secret = load_credentials(profile)
    return bool(api and secret)


class FedExCredentialsDialog(tk.Toplevel):
    WIDTH = 660
    HEIGHT = 700

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.result = False
        self.entries = {}
        self.title("FedEx API Credentials")
        self.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.minsize(self.WIDTH, self.HEIGHT)
        self.resizable(False, False)
        self.configure(bg="#f4f6f9")
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._build()
        self.after(80, self._center)

    def _build(self):
        # ---- header -------------------------------------------------
        header = tk.Frame(self, bg="#0d2137", height=68)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="FedEx API Credentials",
            bg="#0d2137",
            fg="white",
            font=("Segoe UI", 16, "bold"),
        ).pack(side="left", padx=24)

        # ---- button bar (packed BEFORE the form so it can never be
        #      pushed off the bottom of the window) -------------------
        buttons = tk.Frame(self, bg="#eef1f5", height=76)
        buttons.pack(fill="x", side="bottom")
        buttons.pack_propagate(False)

        tk.Button(
            buttons,
            text="Cancel",
            command=self._close,
            bg="#dfe5ec",
            fg="#263342",
            activebackground="#d2d9e2",
            relief="flat",
            font=("Segoe UI", 10),
            width=16,
            cursor="hand2",
        ).pack(side="right", padx=(10, 22), pady=18, ipady=8)

        tk.Button(
            buttons,
            text="Save Changes",
            command=self._save,
            bg="#1a7a3c",
            fg="white",
            activebackground="#146631",
            activeforeground="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            width=20,
            cursor="hand2",
        ).pack(side="right", pady=18, ipady=8)

        # ---- intro text ---------------------------------------------
        tk.Label(
            self,
            text=(
                "Enter the production API Key and Secret Key for UNI and "
                "FENIX. Credentials are stored securely in Windows "
                "Credential Manager."
            ),
            bg="#f4f6f9",
            fg="#4f5d6c",
            justify="left",
            wraplength=self.WIDTH - 60,
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=24, pady=(16, 10), side="top")

        # ---- form ----------------------------------------------------
        form = tk.Frame(
            self,
            bg="white",
            highlightthickness=1,
            highlightbackground="#d5dae2",
        )
        form.pack(fill="both", expand=True, padx=24, pady=(0, 14), side="top")
        form.grid_columnconfigure(1, weight=1)

        row = 0
        for index, profile in enumerate(PROFILES):
            tk.Label(
                form,
                text=f"{profile} FedEx Project",
                bg="white",
                fg="#1b2735",
                font=("Segoe UI", 12, "bold"),
            ).grid(
                row=row,
                column=0,
                columnspan=2,
                sticky="w",
                padx=20,
                pady=(18 if index == 0 else 26, 8),
            )
            row += 1

            api_key, existing_secret = load_credentials(profile)

            tk.Label(
                form,
                text="API Key:",
                bg="white",
                fg="#263342",
                font=("Segoe UI", 9, "bold"),
                width=14,
                anchor="w",
            ).grid(row=row, column=0, sticky="w", padx=(20, 10), pady=5)

            api_var = tk.StringVar(value=api_key)
            tk.Entry(
                form,
                textvariable=api_var,
                font=("Segoe UI", 10),
                bd=1,
                relief="solid",
            ).grid(row=row, column=1, sticky="ew", padx=(0, 20), pady=5, ipady=7)
            row += 1

            tk.Label(
                form,
                text="Secret Key:",
                bg="white",
                fg="#263342",
                font=("Segoe UI", 9, "bold"),
                width=14,
                anchor="w",
            ).grid(row=row, column=0, sticky="w", padx=(20, 10), pady=5)

            secret_var = tk.StringVar()
            tk.Entry(
                form,
                textvariable=secret_var,
                show="\u2022",
                font=("Segoe UI", 10),
                bd=1,
                relief="solid",
            ).grid(row=row, column=1, sticky="ew", padx=(0, 20), pady=5, ipady=7)
            row += 1

            configured = bool(api_key and existing_secret)
            tk.Label(
                form,
                text="Status: " + ("Saved securely" if configured else "Not configured"),
                bg="white",
                fg="#1a7a3c" if configured else "#a02020",
                font=("Segoe UI", 8, "bold"),
            ).grid(row=row, column=1, sticky="w", padx=(0, 20), pady=(0, 2))
            row += 1

            tk.Label(
                form,
                text=(
                    "Leave Secret Key blank to keep the saved secret."
                    if existing_secret
                    else "Enter the Secret Key from the FedEx portal."
                ),
                bg="white",
                fg="#6a7480",
                font=("Segoe UI", 8),
            ).grid(row=row, column=1, sticky="w", padx=(0, 20), pady=(0, 6))
            row += 1

            self.entries[profile] = {
                "api_var": api_var,
                "secret_var": secret_var,
                "has_existing_secret": bool(existing_secret),
            }

        # spacer row so the fields stay top-aligned inside the card
        form.grid_rowconfigure(row, weight=1)

        self.bind("<Return>", lambda _event: self._save())
        self.bind("<Escape>", lambda _event: self._close())

    def _save(self):
        try:
            for profile, fields in self.entries.items():
                old_api, old_secret = load_credentials(profile)
                api = fields["api_var"].get().strip() or old_api
                secret = fields["secret_var"].get().strip() or old_secret
                if api or secret:
                    save_credentials(profile, api, secret)
            self.result = True
            messagebox.showinfo(
                "FedEx credentials",
                "FedEx credentials were saved securely.",
                parent=self,
            )
            self._close()
        except Exception as exc:
            messagebox.showerror(
                "Could not save credentials",
                f"{type(exc).__name__}: {exc}",
                parent=self,
            )

    def _close(self):
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _center(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = self.parent.winfo_rootx() + max(0, (self.parent.winfo_width() - w) // 2)
        y = self.parent.winfo_rooty() + max(0, (self.parent.winfo_height() - h) // 2)
        self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")


def show_credentials_dialog(parent) -> bool:
    dialog = FedExCredentialsDialog(parent)
    parent.wait_window(dialog)
    return dialog.result