import os
import re
import tkinter as tk
from tkinter import messagebox, ttk
import fitz
from PIL import Image, ImageTk

SHEET_OPTIONS = ("UNI", "EMBY", "FENIX", "SOL")
CARRIER_OPTIONS = (
    "UPS GROUND", "UPS NEXT DAY AIR", "UPS 2ND DAY AIR",
    "FEDEX", "BX FX S/O", "BX FX P/O", "USPS",
    "MALCA-AMIT", "BRINKS",
)
FIELDS = (
    ("sheet", "Sheet"),
    ("ship_to", "Send To"),
    ("invoice", "INV / PO Number"),
    ("tracking_number", "Tracking Number"),
    ("carrier", "Carrier"),
)

def _clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()

def missing_required_fields(record):
    return [key for key, _ in FIELDS if not _clean(record.get(key))]

def needs_manual_review(record):
    return bool(missing_required_fields(record))

class ManualReviewDialog(tk.Toplevel):
    def __init__(self, parent, pdf_path, record):
        super().__init__(parent)
        self.parent = parent
        self.pdf_path = pdf_path
        self.original = dict(record or {})
        self.result = None
        self.doc = None
        self.page_index = 0
        self.photo = None
        self.vars = {}
        self.widgets = {}

        self.title("Shipment Label Needs Review")
        self.geometry("1180x720")
        self.minsize(980, 620)
        self.configure(bg="#f4f6f9")
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        self.build_ui()
        self.load_pdf()
        self.after(80, self.focus_missing)

    def build_ui(self):
        header = tk.Frame(self, bg="#0d2137", height=58)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="Shipment label requires manual completion",
            bg="#0d2137", fg="white",
            font=("Segoe UI", 15, "bold"),
        ).pack(side="left", padx=18)

        body = tk.Frame(self, bg="#f4f6f9")
        body.pack(fill="both", expand=True, padx=14, pady=14)
        body.grid_columnconfigure(0, minsize=390)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        left = tk.Frame(body, bg="white", bd=1, relief="solid")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)

        tk.Label(
            left, text="Shipment Details",
            bg="white", fg="#1b2735",
            font=("Segoe UI", 13, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=18, pady=(18, 4))

        tk.Label(
            left,
            text="Complete the yellow fields, verify the others, then click Done.",
            bg="white", fg="#586574",
            wraplength=340, justify="left",
            font=("Segoe UI", 9),
        ).grid(row=1, column=0, sticky="w", padx=18, pady=(0, 14))

        form = tk.Frame(left, bg="white")
        form.grid(row=2, column=0, sticky="ew", padx=18)
        form.grid_columnconfigure(0, weight=1)

        missing = set(missing_required_fields(self.original))

        row = 0
        for key, label in FIELDS:
            tk.Label(
                form, text=label, bg="white",
                fg="#263342", font=("Segoe UI", 9, "bold"),
            ).grid(row=row, column=0, sticky="w", pady=(10 if row else 0, 4))

            var = tk.StringVar(value=_clean(self.original.get(key)))
            self.vars[key] = var

            if key == "sheet":
                widget = ttk.Combobox(
                    form, textvariable=var, values=SHEET_OPTIONS,
                    state="normal", font=("Segoe UI", 10),
                )
            elif key == "carrier":
                widget = ttk.Combobox(
                    form, textvariable=var, values=CARRIER_OPTIONS,
                    state="normal", font=("Segoe UI", 10),
                )
            else:
                widget = tk.Entry(
                    form, textvariable=var,
                    font=("Segoe UI", 10), bd=1, relief="solid",
                    bg="#fff4cc" if key in missing else "white",
                )

            widget.grid(row=row + 1, column=0, sticky="ew", ipady=6)
            self.widgets[key] = widget
            row += 2

        button_frame = tk.Frame(left, bg="white")
        button_frame.grid(row=3, column=0, sticky="ew", padx=18, pady=18)
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        tk.Button(
            button_frame, text="Cancel", command=self.cancel,
            bg="#e5e9ef", relief="flat",
            font=("Segoe UI", 10),
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6), ipady=8)

        tk.Button(
            button_frame, text="Done — Save to Excel",
            command=self.submit,
            bg="#1a7a3c", fg="white", relief="flat",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=1, sticky="ew", padx=(6, 0), ipady=8)

        right = tk.Frame(body, bg="white", bd=1, relief="solid")
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        top = tk.Frame(right, bg="white")
        top.grid(row=0, column=0, sticky="ew", padx=14, pady=10)
        top.grid_columnconfigure(1, weight=1)

        tk.Label(
            top, text="Label Preview",
            bg="white", fg="#1b2735",
            font=("Segoe UI", 13, "bold"),
        ).grid(row=0, column=0, sticky="w")

        self.page_label = tk.Label(
            top, text="", bg="white",
            fg="#667382", font=("Segoe UI", 9),
        )
        self.page_label.grid(row=0, column=1, sticky="e")

        preview_frame = tk.Frame(right, bg="#2b3038")
        preview_frame.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 8))
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        self.preview = tk.Label(
            preview_frame, text="Loading preview...",
            bg="#2b3038", fg="white",
        )
        self.preview.grid(row=0, column=0, sticky="nsew")

        nav = tk.Frame(right, bg="white")
        nav.grid(row=2, column=0, pady=(0, 12))

        self.prev_btn = tk.Button(nav, text="◀ Previous", command=self.previous)
        self.prev_btn.pack(side="left", padx=4)
        self.next_btn = tk.Button(nav, text="Next ▶", command=self.next)
        self.next_btn.pack(side="left", padx=4)

        self.bind("<Return>", lambda _e: self.submit())
        self.bind("<Escape>", lambda _e: self.cancel())

    def load_pdf(self):
        try:
            if not os.path.isfile(self.pdf_path):
                raise FileNotFoundError(self.pdf_path)
            self.doc = fitz.open(self.pdf_path)
            if len(self.doc) == 0:
                raise ValueError("PDF contains no pages")
            self.render_page()
        except Exception as exc:
            self.preview.config(
                text=f"Preview unavailable\n\n{type(exc).__name__}: {exc}",
                image="",
            )
            self.prev_btn.config(state="disabled")
            self.next_btn.config(state="disabled")

    def render_page(self):
        if not self.doc:
            return
        page = self.doc.load_page(self.page_index)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.4, 1.4), alpha=False)
        image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        max_w, max_h = 720, 560
        image.thumbnail((max_w, max_h), Image.LANCZOS)
        self.photo = ImageTk.PhotoImage(image)
        self.preview.config(image=self.photo, text="")
        self.page_label.config(text=f"Page {self.page_index + 1} of {len(self.doc)}")
        self.prev_btn.config(state="normal" if self.page_index > 0 else "disabled")
        self.next_btn.config(
            state="normal" if self.page_index < len(self.doc) - 1 else "disabled"
        )

    def previous(self):
        if self.doc and self.page_index > 0:
            self.page_index -= 1
            self.render_page()

    def next(self):
        if self.doc and self.page_index < len(self.doc) - 1:
            self.page_index += 1
            self.render_page()

    def focus_missing(self):
        missing = missing_required_fields(self.original)
        key = missing[0] if missing else "ship_to"
        widget = self.widgets.get(key)
        if widget:
            widget.focus_set()
            if isinstance(widget, tk.Entry):
                widget.selection_range(0, "end")

    def submit(self):
        corrected = dict(self.original)
        for key, _ in FIELDS:
            corrected[key] = _clean(self.vars[key].get())

        remaining = missing_required_fields(corrected)
        if remaining:
            labels = [label for key, label in FIELDS if key in remaining]
            messagebox.showwarning(
                "Missing shipment details",
                "Complete these fields before saving:\n\n"
                + "\n".join(f"• {label}" for label in labels),
                parent=self,
            )
            self.widgets[remaining[0]].focus_set()
            return

        self.result = corrected
        self.close()

    def cancel(self):
        if messagebox.askyesno(
            "Cancel manual review?",
            "The shipment will not be written to Excel until the missing information is completed.",
            parent=self,
        ):
            self.result = None
            self.close()

    def close(self):
        try:
            if self.doc:
                self.doc.close()
        except Exception:
            pass
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

def show_manual_review(parent, pdf_path, record):
    dialog = ManualReviewDialog(parent, pdf_path, record)
    parent.wait_window(dialog)
    return dialog.result