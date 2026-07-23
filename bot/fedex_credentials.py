"""Secure UNI/FENIX FedEx credentials in Windows Credential Manager."""
from __future__ import annotations
import tkinter as tk
from tkinter import messagebox
import keyring

SERVICE_NAME = "UniCreationShipmentBot.FedEx"
PROFILES = ("UNI", "FENIX")

def _username(profile: str, field: str) -> str:
    profile = profile.strip().upper(); field = field.strip().upper()
    if profile not in PROFILES: raise ValueError(f"Unsupported profile: {profile}")
    if field not in {"API_KEY", "SECRET_KEY"}: raise ValueError(field)
    return f"FEDEX_{profile}_{field}"

def save_credentials(profile: str, api_key: str, secret_key: str) -> None:
    profile = profile.strip().upper(); api_key = api_key.strip(); secret_key = secret_key.strip()
    if not api_key or not secret_key: raise ValueError(f"{profile}: API Key and Secret Key are required.")
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
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent; self.result = False; self.entries = {}
        self.title("FedEx API Credentials"); self.geometry("610x390"); self.resizable(False, False)
        self.configure(bg="#f4f6f9"); self.transient(parent); self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._build(); self.after(80, self._center)

    def _build(self):
        header = tk.Frame(self, bg="#0d2137", height=66)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="FedEx API Credentials",
            bg="#0d2137",
            fg="white",
            font=("Segoe UI", 15, "bold"),
        ).pack(side="left", padx=22)

        # Always-visible footer.
        footer = tk.Frame(
            self,
            bg="#eef1f5",
            height=82,
            highlightthickness=1,
            highlightbackground="#d5dae2",
        )
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)

        tk.Button(
            footer,
            text="Cancel",
            command=self._close,
            bg="#dfe5ec",
            fg="#263342",
            activebackground="#d2d9e2",
            relief="flat",
            font=("Segoe UI", 10),
            width=16,
            cursor="hand2",
        ).pack(
            side="right",
            padx=(10, 20),
            pady=18,
            ipady=8,
        )

        tk.Button(
            footer,
            text="Save Credentials",
            command=self._save,
            bg="#1a7a3c",
            fg="white",
            activebackground="#146631",
            activeforeground="white",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            width=22,
            cursor="hand2",
        ).pack(
            side="right",
            pady=18,
            ipady=8,
        )

        # Scrollable content area.
        outer = tk.Frame(self, bg="#f4f6f9")
        outer.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(
            outer,
            bg="#f4f6f9",
            highlightthickness=0,
            bd=0,
        )
        scrollbar = ttk.Scrollbar(
            outer,
            orient="vertical",
            command=canvas.yview,
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        content = tk.Frame(canvas, bg="#f4f6f9")
        content_window = canvas.create_window(
            (0, 0),
            window=content,
            anchor="nw",
        )

        def update_scroll_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def fit_content_width(event):
            canvas.itemconfigure(
                content_window,
                width=max(event.width, 1),
            )

        content.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", fit_content_width)

        def on_mousewheel(event):
            canvas.yview_scroll(
                int(-1 * (event.delta / 120)),
                "units",
            )

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        tk.Label(
            content,
            text=(
                "Enter the production API Key and Secret Key for each "
                "FedEx project. Credentials are stored securely in "
                "Windows Credential Manager for this Windows user."
            ),
            bg="#f4f6f9",
            fg="#4f5d6c",
            justify="left",
            wraplength=680,
            font=("Segoe UI", 9),
        ).pack(
            anchor="w",
            fill="x",
            padx=24,
            pady=(20, 14),
        )

        form = tk.Frame(
            content,
            bg="white",
            highlightthickness=1,
            highlightbackground="#d5dae2",
        )
        form.pack(
            fill="x",
            expand=False,
            padx=24,
            pady=(0, 24),
        )
        form.grid_columnconfigure(1, weight=1)

        row = 0

        for profile in PROFILES:
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
                pady=(22 if row == 0 else 26, 10),
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
            ).grid(
                row=row,
                column=0,
                sticky="w",
                padx=(20, 10),
                pady=6,
            )

            api_var = tk.StringVar(value=api_key)
            api_entry = tk.Entry(
                form,
                textvariable=api_var,
                font=("Segoe UI", 10),
                bd=1,
                relief="solid",
            )
            api_entry.grid(
                row=row,
                column=1,
                sticky="ew",
                padx=(0, 20),
                pady=6,
                ipady=8,
            )
            row += 1

            tk.Label(
                form,
                text="Secret Key:",
                bg="white",
                fg="#263342",
                font=("Segoe UI", 9, "bold"),
                width=14,
                anchor="w",
            ).grid(
                row=row,
                column=0,
                sticky="w",
                padx=(20, 10),
                pady=6,
            )

            secret_var = tk.StringVar()
            secret_entry = tk.Entry(
                form,
                textvariable=secret_var,
                show="•",
                font=("Segoe UI", 10),
                bd=1,
                relief="solid",
            )
            secret_entry.grid(
                row=row,
                column=1,
                sticky="ew",
                padx=(0, 20),
                pady=6,
                ipady=8,
            )
            row += 1

            status = (
                "Saved securely"
                if api_key and existing_secret
                else "Not configured"
            )

            tk.Label(
                form,
                text=f"Status: {status}",
                bg="white",
                fg="#1a7a3c" if status == "Saved securely" else "#a02020",
                font=("Segoe UI", 8, "bold"),
            ).grid(
                row=row,
                column=1,
                sticky="w",
                padx=(0, 20),
                pady=(1, 2),
            )
            row += 1

            tk.Label(
                form,
                text=(
                    "Leave Secret Key blank to keep the currently saved secret."
                    if existing_secret
                    else "Enter the Secret Key shown in the FedEx portal."
                ),
                bg="white",
                fg="#6a7480",
                font=("Segoe UI", 8),
                wraplength=500,
                justify="left",
            ).grid(
                row=row,
                column=1,
                sticky="w",
                padx=(0, 20),
                pady=(0, 8),
            )
            row += 1

            self.entries[profile] = {
                "api_var": api_var,
                "secret_var": secret_var,
                "has_existing_secret": bool(existing_secret),
            }

        # Give the canvas time to calculate the full scrollable height.
        self.after(50, update_scroll_region)

        self.bind("<Return>", lambda _event: self._save())
        self.bind("<Escape>", lambda _event: self._close())

    def _save(self):
        try:
            for profile, (api_var, secret_var) in self.entries.items():
                old_api, old_secret = load_credentials(profile)
                api = api_var.get().strip() or old_api
                secret = secret_var.get().strip() or old_secret
                if api or secret: save_credentials(profile, api, secret)
            self.result = True
            messagebox.showinfo("FedEx credentials", "FedEx credentials were saved securely.", parent=self)
            self._close()
        except Exception as exc:
            messagebox.showerror("Could not save credentials", f"{type(exc).__name__}: {exc}", parent=self)

    def _close(self):
        try: self.grab_release()
        except Exception: pass
        self.destroy()

    def _center(self):
        self.update_idletasks(); w=self.winfo_width(); h=self.winfo_height()
        x=self.parent.winfo_rootx()+max(0,(self.parent.winfo_width()-w)//2); y=self.parent.winfo_rooty()+max(0,(self.parent.winfo_height()-h)//2)
        self.geometry(f"{w}x{h}+{x}+{y}")

def show_credentials_dialog(parent) -> bool:
    dialog = FedExCredentialsDialog(parent); parent.wait_window(dialog); return dialog.result