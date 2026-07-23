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
        header = tk.Frame(self, bg="#0d2137", height=58); header.pack(fill="x"); header.pack_propagate(False)
        tk.Label(header, text="FedEx API Credentials", bg="#0d2137", fg="white", font=("Segoe UI",14,"bold")).pack(side="left", padx=18)
        tk.Label(self, text="Stored securely in Windows Credential Manager for this Windows user. Nothing is saved in GitHub or config.json.", bg="#f4f6f9", fg="#4f5d6c", wraplength=560, justify="left", font=("Segoe UI",9)).pack(anchor="w", padx=20, pady=(16,10))
        form = tk.Frame(self, bg="#f4f6f9"); form.pack(fill="x", padx=20); form.grid_columnconfigure(1, weight=1)
        row = 0
        for profile in PROFILES:
            api, secret = load_credentials(profile)
            tk.Label(form, text=f"{profile} FedEx Project", bg="#f4f6f9", fg="#1b2735", font=("Segoe UI",11,"bold")).grid(row=row,column=0,columnspan=2,sticky="w",pady=(10 if row else 0,6)); row += 1
            tk.Label(form,text="API Key:",bg="#f4f6f9",width=12,anchor="w").grid(row=row,column=0,sticky="w",pady=3)
            api_var = tk.StringVar(value=api); tk.Entry(form,textvariable=api_var,width=55,font=("Segoe UI",9)).grid(row=row,column=1,sticky="ew",pady=3); row += 1
            tk.Label(form,text="Secret Key:",bg="#f4f6f9",width=12,anchor="w").grid(row=row,column=0,sticky="w",pady=3)
            secret_var = tk.StringVar(); tk.Entry(form,textvariable=secret_var,width=55,show="•",font=("Segoe UI",9)).grid(row=row,column=1,sticky="ew",pady=3); row += 1
            status = "Saved" if api and secret else "Not configured"
            tk.Label(form,text=f"Current status: {status}",bg="#f4f6f9",fg="#1a7a3c" if status=="Saved" else "#a02020",font=("Segoe UI",8)).grid(row=row,column=1,sticky="w",pady=(0,5)); row += 1
            self.entries[profile] = (api_var, secret_var)
        buttons = tk.Frame(self,bg="#f4f6f9"); buttons.pack(fill="x",padx=20,pady=18)
        tk.Button(buttons,text="Cancel",command=self._close,bg="#dde3ec",relief="flat",width=14).pack(side="right",padx=(8,0))
        tk.Button(buttons,text="Save Securely",command=self._save,bg="#1a7a3c",fg="white",relief="flat",font=("Segoe UI",9,"bold"),width=18).pack(side="right")

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