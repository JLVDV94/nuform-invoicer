# -*- coding: utf-8 -*-
"""
Created on Fri Aug  1 09:02:13 2025

@author: Acc
"""

# =========================
# NuForm Invoicer (Track A updates ready) - Windows-first
# =========================
# Notes:
# - Replace UPDATE_MANIFEST_URL with your hosted latest.json URL.
# - Build with PyInstaller; wrap with Inno Setup for an installer.
# - App reads/writes under ~/Documents/NuForm Invoicing by default.
#
# After you publish an update:
#   1) Bump __version__ below, rebuild EXE + installer
#   2) Upload installer to your hosting
#   3) Update latest.json with new version, notes, windows_installer_url, windows_sha256
#   4) Users click "Check for updates" in the app (or you can auto-check on startup)

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from fpdf import FPDF
import pandas as pd
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import os
import platform
import subprocess
import json
from datetime import datetime
from typing import Dict, List, Optional
from pathlib import Path
import urllib.request
import tempfile
import hashlib

# =========================
# App Version & Updates
# =========================
__version__ = "1.0.0"  # << bump each release
UPDATE_MANIFEST_URL = "https://jlvdv94.github.io/nuform-invoicer/latest.json"
 # <-- change to your URL

def _parse_version(v: str):
    return tuple(int(x) for x in str(v).strip().split("."))

def _sha256_of(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()

def check_for_updates_windows(root_widget: Optional[tk.Tk] = None):
    """
    Track A update flow:
      - Fetch manifest (latest.json)
      - Compare version
      - Download installer
      - Verify SHA256 (if provided)
      - Launch installer (silent by default)
      - Close the running app
    """
    try:
        with urllib.request.urlopen(UPDATE_MANIFEST_URL, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        messagebox.showerror("Update", f"Could not reach update server.\n\n{e}")
        return

    latest = (data.get("version") or "").strip()
    notes = data.get("notes", "")
    win_url = (data.get("windows_installer_url") or "").strip()
    win_sha = (data.get("windows_sha256") or "").strip().lower()

    if not latest or not win_url:
        messagebox.showerror("Update", "Update manifest is missing required fields.")
        return

    try:
        if _parse_version(latest) <= _parse_version(__version__):
            messagebox.showinfo("Up to date", f"You are on the latest version ({__version__}).")
            return
    except Exception:
        # If parsing fails, continue to offer update (safer)
        pass

    msg = f"Version {latest} is available.\n\nRelease notes:\n{notes}\n\nDownload & install now?"
    if not messagebox.askyesno("Update available", msg):
        return

    # Download installer
    try:
        dl_path = os.path.join(tempfile.gettempdir(), os.path.basename(win_url))
        urllib.request.urlretrieve(win_url, dl_path)
    except Exception as e:
        messagebox.showerror("Update", f"Download failed.\n\n{e}")
        return

    # Verify integrity if checksum provided
    if win_sha:
        try:
            actual = _sha256_of(dl_path)
            if actual != win_sha:
                messagebox.showerror("Update", "Downloaded file failed integrity check.")
                try:
                    os.remove(dl_path)
                except Exception:
                    pass
                return
        except Exception as e:
            messagebox.showerror("Update", f"Integrity check error.\n\n{e}")
            return

    # Launch installer (silent flags recommended for Inno Setup)
    try:
        subprocess.Popen([dl_path, "/VERYSILENT", "/NORESTART"], shell=False)
        if root_widget is not None:
            root_widget.destroy()
    except Exception as e:
        messagebox.showerror("Update", f"Could not start installer.\n\n{e}")


# =========================
# Configuration & Constants
# =========================
# Portable, user-friendly data directory:
APP_DIR = Path.home() / "Documents" / "NuForm Invoicing"
APP_DIR.mkdir(parents=True, exist_ok=True)

# If you already have files elsewhere, place copies into this folder once.
SERVICES_FILE = str(APP_DIR / "Services and Prices.xlsx")
ICD10_PRIMARY_FILE = str(APP_DIR / "icd10_codes.xlsx")
ICD10_SECONDARY_FILE = str(APP_DIR / "Secondary ICD10 Codes.xlsx")
PATIENTS_FILE = str(APP_DIR / "patients.csv")
LOGO_PATH = str(APP_DIR / "NuForm Health 2.png")
INVOICE_COUNTER_FILE = str(APP_DIR / "invoice_counter.json")

CURRENCY_PREFIX = "R"
DEFAULT_VAT_RATE = Decimal("0.15")  # 15%
DATE_FMT = "%Y-%m-%d"

# Expected column fallbacks
SERVICE_COL = "Service"
TARIFF_COL_CANDIDATES = ["Tariff Code", "Tarrif Code", "Tariff", "Tarrif"]
NAPPI_COL_CANDIDATES = ["NAPPI Code", "NAPPI"]
PRICE_COL_CANDIDATES = ["Price", "Unit Price"]

ICD10_CODE_COL = "Code"
ICD10_DESC_COL = "Description"

# Patient fields (CSV schema)
PATIENT_FIELDS = [
    "Patient File No", "Name", "Surname", "ID", "Address",
    "Phone", "Email", "Medical Aid", "Medical Aid Plan", "Membership No"
]

# Optional calendar support
try:
    from tkcalendar import DateEntry
    TKCAL_AVAILABLE = True
except Exception:
    TKCAL_AVAILABLE = False

# =========================
# Helpers
# =========================

def money_to_decimal(s: str) -> Decimal:
    if s is None:
        raise InvalidOperation
    clean = str(s).strip().replace(CURRENCY_PREFIX, "").replace(",", "")
    if clean == "":
        raise InvalidOperation
    return Decimal(clean)

def fmt_money(x: Decimal) -> str:
    return f"{CURRENCY_PREFIX}{x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)}"

def load_excel_safe(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path)
    except Exception:
        return pd.DataFrame()

def find_first_existing_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def normalize_long_number(s: str) -> str:
    """
    Convert '9.40215E+13' etc. to plain digits. If conversion fails, return the original.
    Also strips non-digits at the end so IDs stay clean.
    """
    if s is None:
        return ""
    raw = str(s).strip()
    try:
        if "e" in raw.lower() or "." in raw:
            d = Decimal(raw)
            raw = format(d.quantize(0), "f")
    except Exception:
        pass
    digits = "".join(ch for ch in raw if ch.isdigit())
    return digits or raw

def ensure_patients_file():
    if not os.path.exists(PATIENTS_FILE):
        df = pd.DataFrame(columns=PATIENT_FIELDS)
        df.to_csv(PATIENTS_FILE, index=False, encoding="utf-8")

def load_patients_df() -> pd.DataFrame:
    ensure_patients_file()
    try:
        df = pd.read_csv(PATIENTS_FILE, dtype=str, encoding="utf-8")
    except Exception:
        df = pd.DataFrame(columns=PATIENT_FIELDS)
    for col in PATIENT_FIELDS:
        if col not in df.columns:
            df[col] = ""
    return df[PATIENT_FIELDS].fillna("")

def upsert_patient_row(new_row: Dict[str, str]):
    df = load_patients_df()

    # normalize all fields to strings and fix ID
    for k in PATIENT_FIELDS:
        v = str(new_row.get(k, "") or "")
        if k == "ID":
            v = normalize_long_number(v)
        new_row[k] = v

    idx = None
    if new_row.get("ID"):
        matches = df.index[df["ID"].astype(str).str.strip().str.lower() == new_row["ID"].strip().lower()].tolist()
        if matches:
            idx = matches[0]
    if idx is None and new_row.get("Patient File No"):
        matches = df.index[df["Patient File No"].astype(str).str.strip().str.lower() == new_row["Patient File No"].strip().lower()].tolist()
        if matches:
            idx = matches[0]

    if idx is None:
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    else:
        for k in PATIENT_FIELDS:
            if new_row.get(k, "") != "":
                df.at[idx, k] = new_row[k]

    df.to_csv(PATIENTS_FILE, index=False, encoding="utf-8")

# ---------- Invoice counter (peek vs persist) ----------

def load_invoice_counter() -> int:
    try:
        with open(INVOICE_COUNTER_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return int(data.get("last", 1000))
    except Exception:
        return 1000

def save_invoice_counter(n: int):
    try:
        with open(INVOICE_COUNTER_FILE, "w", encoding="utf-8") as f:
            json.dump({"last": int(n)}, f)
    except Exception:
        pass

def peek_next_invoice_number() -> int:
    return load_invoice_counter() + 1

def format_invoice_no(n: int) -> str:
    return f"INV-{n}"

def persist_invoice_if_new(inv_string: str):
    """Persist counter to this invoice number if it's newer than stored."""
    try:
        n = int(str(inv_string).replace("INV-", "").strip())
    except Exception:
        return
    last = load_invoice_counter()
    if n > last:
        save_invoice_counter(n)

def parse_date_safe(s: str) -> datetime:
    try:
        return datetime.strptime(s, DATE_FMT)
    except Exception:
        return datetime.today()

def widget_get_date_str(w) -> str:
    if TKCAL_AVAILABLE and isinstance(w, DateEntry):
        return w.get_date().strftime(DATE_FMT)
    else:
        val = (w.get() if hasattr(w, "get") else "").strip()
        try:
            datetime.strptime(val, DATE_FMT)
            return val
        except Exception:
            return datetime.today().strftime(DATE_FMT)

def widget_set_date(w, date_str: str):
    if TKCAL_AVAILABLE and isinstance(w, DateEntry):
        try:
            w.set_date(parse_date_safe(date_str))
        except Exception:
            w.set_date(datetime.today())
    else:
        try:
            if hasattr(w, "delete"):
                w.delete(0, tk.END)
                w.insert(0, parse_date_safe(date_str).strftime(DATE_FMT))
        except Exception:
            pass

# =========================
# Data Loading
# =========================

services_df = load_excel_safe(SERVICES_FILE)
icd10_df = load_excel_safe(ICD10_PRIMARY_FILE)
icd10_sec_df = load_excel_safe(ICD10_SECONDARY_FILE)

service_ok = SERVICE_COL in services_df.columns
tariff_col = find_first_existing_column(services_df, TARIFF_COL_CANDIDATES) or ""
nappi_col = find_first_existing_column(services_df, NAPPI_COL_CANDIDATES) or ""
price_col = find_first_existing_column(services_df, PRICE_COL_CANDIDATES) or ""

icd10_ok = ICD10_CODE_COL in icd10_df.columns and ICD10_DESC_COL in icd10_df.columns
icd10_sec_ok = ICD10_CODE_COL in icd10_sec_df.columns and ICD10_DESC_COL in icd10_sec_df.columns

# =========================
# PDF
# =========================

class InvoicePDF(FPDF):
    def __init__(self, *args, **kwargs):
        self.is_compact = kwargs.pop("is_compact", False)
        super().__init__(*args, **kwargs)

    def header(self):
        pw = self.w
        if self.is_compact:
            self.set_auto_page_break(auto=True, margin=8)
            logo_w = 32
            logo_x = (pw - logo_w) / 2
            logo_y = 4
            if os.path.exists(LOGO_PATH):
                self.image(LOGO_PATH, logo_x, logo_y, logo_w)
            safe_y = int(logo_y + logo_w + 2)
            self.set_y(safe_y)
            self.set_font("Helvetica", '', 8.2)
            self.cell(0, 4.2, "Dr Melissa Janse van Vuren - Sport & Lifestyle Medicine Practitioner", ln=True, align='C')
            self.cell(0, 4.2, "MBChB, MSc (Sport & Exercise Medicine), HPCSA: MP0820040, Practice No: 0824593", ln=True, align='C')
            self.set_font("Helvetica", '', 7.8)
            self.multi_cell(0, 3.8, "Address: 6 Trappes Street, Langerug, Worcester, 6850\nPhone: 068 948 1808 | Email: dr.melissa@nuformhealth.co.za", align='C')
            self.ln(0.5)
        else:
            self.set_auto_page_break(auto=True, margin=12)
            logo_w = 40
            logo_x = (pw - logo_w) / 2
            logo_y = 6
            if os.path.exists(LOGO_PATH):
                self.image(LOGO_PATH, logo_x, logo_y, logo_w)
            safe_y = int(logo_y + logo_w + 2)
            self.set_y(safe_y)
            self.set_font("Helvetica", '', 9.4)
            self.cell(0, 5, "Dr Melissa Janse van Vuren - Sport & Lifestyle Medicine Practitioner", ln=True, align='C')
            self.cell(0, 5, "MBChB, MSc (Sport & Exercise Medicine), HPCSA: MP0820040, Practice No: 0824593", ln=True, align='C')
            self.set_font("Helvetica", '', 8.8)
            self.multi_cell(0, 4.6, "Address: 6 Trappes Street, Langerug, Worcester, 6850\nPhone: 068 948 1808 | Email: dr.melissa@nuformhealth.co.za", align='C')
            self.ln(1)

    def footer(self):
        if self.is_compact:
            self.set_y(-16)
            self.set_font("Helvetica", 'I', 7.3)
            self.multi_cell(0, 3.8,
                "Payment Terms: Due on receipt.\n"
                "Bank: FNB | Acc: 63167859273 | Branch: 250655 | Ref: Patient Name",
                align='C'
            )
            self.set_y(-7)
            self.cell(0, 5, f"Page {self.page_no()}", align='C')
        else:
            self.set_y(-20)
            self.set_font("Helvetica", 'I', 7.8)
            self.multi_cell(0, 4.2,
                "Payment Terms: Due on receipt.\n"
                "Bank details: NuForm Health (Pty) Ltd | Acc No: 63167859273 | Bank: FNB | Account Type: Cheque | Branch Code: 250655 | Ref: Patient Name\n"
                "Thank you for your trust in our care.",
                align='C'
            )
            self.set_y(-8)
            self.cell(0, 6, f"Page {self.page_no()}", align='C')

# =========================
# Autocomplete Combobox
# =========================

class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, master=None, values_list=None, **kwargs):
        super().__init__(master, **kwargs)
        self._all_values = sorted(set(values_list or []), key=lambda s: s.lower() if isinstance(s, str) else "")
        self.configure(values=self._all_values)
        self.bind("<KeyRelease>", self._on_keyrelease)

    def set_all_values(self, values_list):
        self._all_values = sorted(set(values_list or []), key=lambda s: s.lower() if isinstance(s, str) else "")
        self.configure(values=self._all_values)

    def _on_keyrelease(self, event):
        txt = self.get().strip().lower()
        if not txt:
            self.configure(values=self._all_values)
            return
        filtered = [v for v in self._all_values if txt in str(v).lower()]
        self.configure(values=filtered)

# =========================
# Selection dialog for multiple patient matches
# =========================

class PatientSelectDialog(tk.Toplevel):
    def __init__(self, parent, matches_df: pd.DataFrame):
        super().__init__(parent)
        self.title("Select Patient")
        self.geometry("700x360")
        self.resizable(True, True)
        self.result = None

        ttk.Label(self, text="Multiple matches found. Double-click to select:").pack(anchor="w", padx=10, pady=6)

        columns = ("Patient File No", "Surname", "Name", "ID", "Phone", "Email")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=10)
        for c in columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120 if c != "Email" else 160, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=6)

        for _, row in matches_df.iterrows():
            self.tree.insert("", "end", values=(
                row.get("Patient File No",""),
                row.get("Surname",""),
                row.get("Name",""),
                row.get("ID",""),
                row.get("Phone",""),
                row.get("Email","")
            ))

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=8)
        ttk.Button(btns, text="Select", command=self._select).pack(side="right", padx=5)
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="right")

        self.tree.bind("<Double-1>", lambda e: self._select())

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.wait_visibility()
        self.focus()
        self.tree.focus_set()

    def _select(self):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel[0], "values")
        self.result = {
            "Patient File No": values[0],
            "Surname": values[1],
            "Name": values[2],
            "ID": values[3],
            "Phone": values[4],
            "Email": values[5],
        }
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()

# =========================
# Tk App
# =========================

class InvoiceApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"NuForm Health Invoice Generator (v{__version__})")

        self.vat_rate: Decimal = DEFAULT_VAT_RATE
        self.last_pdf_path: Optional[str] = None
        self.lines: List[Dict[str, tk.Widget]] = []
        self.ui_ready: bool = False

        self.patients_df = load_patients_df()
        self._refresh_patient_autocomplete_sources()

        self.entries: Dict[str, ttk.Entry] = {}
        labels = [
            "Invoice No",
            "Invoice Date",
            "Patient File No",
            "Name",
            "Surname",
            "ID",
            "Address",
            "Phone",
            "Email",
            "Medical Aid",
            "Medical Aid Plan",
            "Membership No"
        ]

        for i, label in enumerate(labels):
            ttk.Label(root, text=label + ":").grid(row=i, column=0, sticky='e', padx=6, pady=2)
            if label == "Invoice No":
                entry = ttk.Entry(root, width=44, state="readonly")
                entry.grid(row=i, column=1, columnspan=3, padx=6, pady=2, sticky='w')
                self.entries[label] = entry
                continue

            if label == "Invoice Date":
                if TKCAL_AVAILABLE:
                    entry = DateEntry(root, width=12, date_pattern="yyyy-mm-dd")
                    entry.set_date(datetime.today())
                    entry.grid(row=i, column=1, padx=6, pady=2, sticky='w')
                    ttk.Label(root, text="").grid(row=i, column=2)
                    ttk.Label(root, text="").grid(row=i, column=3)
                else:
                    entry = ttk.Entry(root, width=44)
                    entry.insert(0, datetime.today().strftime(DATE_FMT))
                    entry.grid(row=i, column=1, columnspan=3, padx=6, pady=2, sticky='w')
                self.entries[label] = entry
                continue

            if label in ("Patient File No", "Name", "Surname", "ID"):
                entry = AutocompleteCombobox(root, values_list=self.patient_value_index[label], width=42, state="normal")
                entry.bind("<<ComboboxSelected>>", self._on_patient_lookup_select)
                entry.bind("<Return>", self._on_patient_lookup_enter)
            else:
                entry = ttk.Entry(root, width=44)
            entry.grid(row=i, column=1, columnspan=3, padx=6, pady=2, sticky='w')
            self.entries[label] = entry

        # peek (don't persist) next invoice no
        self._assign_peek_invoice_number()

        # Actions row
        actions_frame = ttk.Frame(root)
        actions_frame.grid(row=len(labels), column=0, columnspan=6, sticky='w', padx=6, pady=(2, 0))
        ttk.Button(actions_frame, text="Load Patient", command=self._lookup_and_fill_patient).grid(row=0, column=0, sticky='w')
        ttk.Button(actions_frame, text="Copy invoice date to all line dates", command=self.copy_invoice_date_to_lines).grid(row=0, column=1, padx=10, sticky='w')
        self.compact_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(actions_frame, text="Compact mode (PDF)", variable=self.compact_var).grid(row=0, column=2, padx=10, sticky='w')
        self.portrait_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(actions_frame, text="Portrait (PDF)", variable=self.portrait_var).grid(row=0, column=3, padx=10, sticky='w')

        # VAT + Minimum Portrait Font controls
        vat_frame = ttk.Frame(root)
        vat_frame.grid(row=len(labels) + 1, column=0, columnspan=6, sticky='w', padx=6, pady=(6, 0))
        ttk.Label(vat_frame, text="VAT rate (%):").grid(row=0, column=0, sticky='w')
        self.vat_var = tk.StringVar(value=f"{(self.vat_rate * 100):.0f}")
        vat_entry = ttk.Entry(vat_frame, textvariable=self.vat_var, width=5)
        vat_entry.grid(row=0, column=1, padx=(4, 12))
        ttk.Label(vat_frame, text="(Prices include VAT)").grid(row=0, column=2, sticky='w')
        vat_entry.bind("<FocusOut>", self._on_vat_change)
        vat_entry.bind("<Return>", self._on_vat_change)

        ttk.Label(vat_frame, text="Min Portrait Font Size:").grid(row=1, column=0, sticky='w', pady=(4,0))
        self.min_portrait_font_var = tk.DoubleVar(value=6.2)
        try:
            min_font_spin = ttk.Spinbox(vat_frame, from_=5.0, to=10.0, increment=0.1,
                                        textvariable=self.min_portrait_font_var, width=6)
        except Exception:
            min_font_spin = tk.Spinbox(vat_frame, from_=5.0, to=10.0, increment=0.1,
                                       textvariable=self.min_portrait_font_var, width=6)
        min_font_spin.grid(row=1, column=1, padx=(4, 12), sticky='w')
        ttk.Label(vat_frame, text="(Portrait table won't shrink below this)").grid(row=1, column=2, sticky='w', pady=(4,0))

        # Table
        self.table_frame = ttk.Frame(root)
        self.table_frame.grid(row=len(labels) + 2, column=0, columnspan=6, pady=8, padx=6, sticky='w')
        headers = ["Date", "Service", "ICD-10 Code", "ICD-10 Secondary", "Tariff", "NAPPI", "Qty", "Price", "Total", ""]
        for i, h in enumerate(headers):
            ttk.Label(self.table_frame, text=h).grid(row=0, column=i, padx=2)

        # Totals
        totals_frame = ttk.Frame(root)
        totals_frame.grid(row=len(labels) + 3, column=0, columnspan=6, sticky='e', padx=6)
        self.ex_lbl = ttk.Label(totals_frame, text=f"Excl. VAT: {fmt_money(Decimal('0'))}")
        self.vat_lbl = ttk.Label(totals_frame, text=f"VAT @ {int(self.vat_rate*100)}%: {fmt_money(Decimal('0'))}")
        self.total_lbl = ttk.Label(totals_frame, text=f"Total (Incl. VAT): {fmt_money(Decimal('0'))}")
        self.ex_lbl.grid(row=0, column=0, padx=8, pady=2)
        self.vat_lbl.grid(row=0, column=1, padx=8, pady=2)
        self.total_lbl.grid(row=0, column=2, padx=8, pady=2)

        # Buttons
        btn_frame = ttk.Frame(root)
        btn_frame.grid(row=len(labels) + 4, column=0, columnspan=6, pady=8)
        ttk.Button(btn_frame, text="Add Line", command=self.add_line_item).grid(row=0, column=0, padx=4)
        ttk.Button(btn_frame, text="Save Draft", command=self.save_draft).grid(row=0, column=1, padx=4)
        ttk.Button(btn_frame, text="Load Draft", command=self.load_draft).grid(row=0, column=2, padx=4)
        ttk.Button(btn_frame, text="New Invoice", command=self.reset_form).grid(row=0, column=3, padx=4)
        ttk.Button(btn_frame, text="Generate PDF", command=self.generate_pdf).grid(row=0, column=4, padx=4)
        ttk.Button(btn_frame, text="Print", command=self.print_pdf).grid(row=0, column=5, padx=4)
        ttk.Button(btn_frame, text="Check for updates", command=self._check_updates).grid(row=0, column=6, padx=4)

        for c in range(6):
            root.grid_columnconfigure(c, weight=0)

        self.ui_ready = True
        self.add_line_item()
        self.recalc_totals()

        # (Optional) Auto-check for updates a bit after startup
        # self.root.after(2500, self._check_updates)

    # ----- Updates -----
    def _check_updates(self):
        if platform.system() == "Windows":
            check_for_updates_windows(self.root)
        else:
            messagebox.showinfo("Updates", "Windows auto-updater is configured. macOS updater will be added next.")

    # ----- Patient lookup -----

    def _assign_peek_invoice_number(self):
        inv = format_invoice_no(peek_next_invoice_number())
        self.entries["Invoice No"].configure(state="normal")
        self.entries["Invoice No"].delete(0, tk.END)
        self.entries["Invoice No"].insert(0, inv)
        self.entries["Invoice No"].configure(state="readonly")

    def _refresh_patient_autocomplete_sources(self):
        df = load_patients_df()
        self.patient_value_index = {
            "Patient File No": sorted(df["Patient File No"].dropna().astype(str).unique()),
            "Name": sorted(df["Name"].dropna().astype(str).unique()),
            "Surname": sorted(df["Surname"].dropna().astype(str).unique()),
            "ID": sorted(df["ID"].dropna().astype(str).unique()),
        }

    def _get_patient_matches(self, key_field: str, key_value: str) -> pd.DataFrame:
        df = load_patients_df()
        if not key_value:
            return pd.DataFrame(columns=PATIENT_FIELDS)
        key_value = str(key_value).strip().lower()
        if key_field not in df.columns:
            return pd.DataFrame(columns=PATIENT_FIELDS)
        mask = df[key_field].astype(str).str.strip().str.lower() == key_value
        if not mask.any():
            mask = df[key_field].astype(str).str.strip().str.lower().str.contains(key_value)
        return df[mask]

    def _fill_patient_row(self, row: pd.Series):
        for col in PATIENT_FIELDS:
            if col in self.entries:
                self.entries[col].delete(0, tk.END)
                self.entries[col].insert(0, str(row.get(col, "") or ""))

    def _select_from_matches(self, matches: pd.DataFrame) -> Optional[pd.Series]:
        dlg = PatientSelectDialog(self.root, matches)
        self.root.wait_window(dlg)
        if dlg.result is None:
            return None
        sel = matches
        for k, v in dlg.result.items():
            if k in sel.columns and v:
                sel = sel[sel[k].astype(str) == v]
        if len(sel) >= 1:
            return sel.iloc[0]
        return matches.iloc[0]

    def _on_patient_lookup_select(self, event=None):
        self._lookup_and_fill_patient()

    def _on_patient_lookup_enter(self, event=None):
        self._lookup_and_fill_patient()

    def _lookup_and_fill_patient(self):
        for key in ["ID", "Patient File No", "Surname", "Name"]:
            val = self.entries[key].get().strip()
            if val:
                matches = self._get_patient_matches(key, val)
                if len(matches) == 1:
                    self._fill_patient_row(matches.iloc[0])
                    return
                elif len(matches) > 1:
                    chosen = self._select_from_matches(matches)
                    if chosen is not None:
                        self._fill_patient_row(chosen)
                        return
        messagebox.showinfo("Not found", "No matching patient found. Complete details and they will be saved for next time.")

    def _get_invoice_date(self) -> str:
        w = self.entries["Invoice Date"]
        return widget_get_date_str(w)

    # ----- Utility buttons -----

    def copy_invoice_date_to_lines(self):
        inv_date = self._get_invoice_date()
        for line in self.lines:
            if "Date" in line:
                widget_set_date(line["Date"], inv_date)

    # ----- Line items -----

    def add_line_item(self, preset: Optional[Dict[str, str]] = None):
        row = len(self.lines) + 1
        fields: Dict[str, tk.Widget] = {}

        if TKCAL_AVAILABLE:
            date_widget = DateEntry(self.table_frame, width=10, date_pattern="yyyy-mm-dd")
            date_widget.set_date(parse_date_safe(self._get_invoice_date()))
        else:
            date_widget = ttk.Entry(self.table_frame, width=10)
            date_widget.insert(0, self._get_invoice_date())
        fields['Date'] = date_widget
        fields['Date'].grid(row=row, column=0, padx=2)

        services = services_df[SERVICE_COL].dropna().astype(str).tolist() if service_ok else []
        fields['Service'] = ttk.Combobox(self.table_frame, width=20, values=services, state="readonly")
        fields['Service'].grid(row=row, column=1, padx=2)
        fields['Service'].bind('<<ComboboxSelected>>', lambda e, f=fields: self.populate_service_data(f))

        if icd10_ok:
            icd10_values = [f"{code} - {desc}" for code, desc in zip(icd10_df[ICD10_CODE_COL], icd10_df[ICD10_DESC_COL])]
        else:
            icd10_values = []
        fields['ICD-10 Code'] = ttk.Combobox(self.table_frame, width=35, values=icd10_values)
        fields['ICD-10 Code'].grid(row=row, column=2, padx=2)

        if icd10_sec_ok:
            icd10_sec_values = [f"{code} - {desc}" for code, desc in zip(icd10_sec_df[ICD10_CODE_COL], icd10_sec_df[ICD10_DESC_COL])]
        else:
            icd10_sec_values = []
        fields['ICD-10 Secondary'] = ttk.Combobox(self.table_frame, width=35, values=icd10_sec_values)
        fields['ICD-10 Secondary'].grid(row=row, column=3, padx=2)

        fields['Tariff'] = ttk.Entry(self.table_frame, width=8)
        fields['Tariff'].grid(row=row, column=4, padx=2)

        fields['NAPPI'] = ttk.Entry(self.table_frame, width=10)
        fields['NAPPI'].grid(row=row, column=5, padx=2)

        fields['Qty'] = ttk.Entry(self.table_frame, width=5)
        fields['Qty'].insert(0, '1')
        fields['Qty'].grid(row=row, column=6, padx=2)

        fields['Price'] = ttk.Entry(self.table_frame, width=10)  # VAT-inclusive
        fields['Price'].grid(row=row, column=7, padx=2)

        fields['Total'] = ttk.Entry(self.table_frame, width=10, state='readonly')
        fields['Total'].grid(row=row, column=8, padx=2)

        remove_btn = ttk.Button(self.table_frame, text="✕", width=3, command=lambda f=fields: self.remove_line_item(f))
        remove_btn.grid(row=row, column=9, padx=2)
        fields['_remove_btn'] = remove_btn

        fields['Qty'].bind('<KeyRelease>', lambda e, f=fields: self.update_total(f))
        fields['Price'].bind('<KeyRelease>', lambda e, f=fields: self.update_total(f))

        self.lines.append(fields)

        if preset:
            for key, val in preset.items():
                if key not in fields:
                    continue
                w = fields[key]
                if key == "Date":
                    widget_set_date(w, val)
                else:
                    if isinstance(w, (ttk.Combobox, ttk.Entry)):
                        w.delete(0, tk.END)
                        w.insert(0, val)

        if self.ui_ready:
            self.update_total(fields)

    def remove_line_item(self, fields: Dict[str, tk.Widget]):
        for w in fields.values():
            try:
                w.destroy()
            except Exception:
                pass
        self.lines = [l for l in self.lines if l is not fields]
        for i, line in enumerate(self.lines, start=1):
            for j, key in enumerate(["Date", "Service", "ICD-10 Code", "ICD-10 Secondary", "Tariff", "NAPPI", "Qty", "Price", "Total"]):
                widget = line.get(key)
                if widget:
                    widget.grid(row=i, column=j)
            if line.get("_remove_btn"):
                line["_remove_btn"].grid(row=i, column=9)
        self.recalc_totals()

    def populate_service_data(self, fields: Dict[str, tk.Widget]):
        service = fields['Service'].get()
        if not (service_ok and price_col):
            return
        row = services_df[services_df[SERVICE_COL] == service]
        if row.empty:
            return
        r = row.iloc[0]

        if tariff_col and tariff_col in r:
            val = "" if pd.isna(r[tariff_col]) else str(r[tariff_col])
            fields['Tariff'].delete(0, tk.END); fields['Tariff'].insert(0, val)

        if nappi_col and nappi_col in r:
            val = "" if pd.isna(r[nappi_col]) else str(r[nappi_col])
            fields['NAPPI'].delete(0, tk.END); fields['NAPPI'].insert(0, val)

        price_raw = "" if pd.isna(r[price_col]) else str(r[price_col])
        try:
            price_dec = money_to_decimal(price_raw)
            fields['Price'].delete(0, tk.END); fields['Price'].insert(0, f"{price_dec}")
        except InvalidOperation:
            pass

        self.update_total(fields)

    def update_total(self, fields: Dict[str, tk.Widget]):
        try:
            qty = money_to_decimal(fields['Qty'].get())
            price_incl = money_to_decimal(fields['Price'].get())
            total = (qty * price_incl).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        except InvalidOperation:
            total = Decimal('0.00')

        fields['Total'].configure(state='normal')
        fields['Total'].delete(0, tk.END)
        fields['Total'].insert(0, f"{total}")
        fields['Total'].configure(state='readonly')

        if self.ui_ready:
            self.recalc_totals()

    def recalc_totals(self):
        if not self.ui_ready:
            return

        gross = Decimal('0.00')  # VAT-inclusive
        for line in self.lines:
            try:
                gross += money_to_decimal(line['Total'].get())
            except InvalidOperation:
                pass

        if self.vat_rate == 0:
            ex = gross
            vat = Decimal('0.00')
        else:
            ex = (gross / (Decimal(1) + self.vat_rate)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            vat = (gross - ex).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

        self.ex_lbl.config(text=f"Excl. VAT: {fmt_money(ex)}")
        self.vat_lbl.config(text=f"VAT @ {int(self.vat_rate*100)}%: {fmt_money(vat)}")
        self.total_lbl.config(text=f"Total (Incl. VAT): {fmt_money(gross)}")

    def _on_vat_change(self, event=None):
        raw = self.vat_var.get().strip().replace("%", "")
        try:
            pct = Decimal(raw)
            self.vat_rate = (pct / Decimal(100)).quantize(Decimal('0.0001'))
        except InvalidOperation:
            messagebox.showerror("Invalid VAT", "Please enter a valid VAT percentage (e.g., 15).")
            self.vat_var.set(f"{(DEFAULT_VAT_RATE*100):.0f}")
            self.vat_rate = DEFAULT_VAT_RATE
        self.recalc_totals()

    # ----- PDF generation -----

    def generate_pdf(self):
        name = self.entries.get("Name").get().strip()
        surname = self.entries.get("Surname").get().strip()
        if not name or not surname:
            messagebox.showerror("Missing Details", "Please enter both Name and Surname before generating the PDF.")
            return

        # Save/refresh patient DB
        patient_row = {col: self.entries[col].get().strip() for col in PATIENT_FIELDS}
        upsert_patient_row(patient_row)
        self.patients_df = load_patients_df()
        self._refresh_patient_autocomplete_sources()
        for key in ["Patient File No", "Name", "Surname", "ID"]:
            w = self.entries[key]
            if isinstance(w, AutocompleteCombobox):
                w.set_all_values(self.patient_value_index[key])

        invoice_no = self.entries["Invoice No"].get().strip()
        invoice_date = self._get_invoice_date()
        portrait = bool(self.portrait_var.get())
        compact = bool(self.compact_var.get())

        # PDF config
        pdf = InvoicePDF(orientation=('P' if portrait else 'L'), unit='mm', format='A4', is_compact=compact)
        if compact:
            pdf.set_margins(6, 6, 6)
        else:
            pdf.set_margins(10, 10, 10)
        pdf.add_page()

        # Tiny spacer after header
        pdf.set_font("Helvetica", 'B', 11 if compact else 12)
        pdf.cell(0, 3 if compact else 4, "", ln=True)

        # ---- Helpers for full-width centered lines
        def full_width(pdf_obj) -> float:
            return pdf_obj.w - pdf_obj.l_margin - pdf_obj.r_margin

        def center_cell(text: str, h: float, font_style: str = '', font_size: float = 9, ln: bool = True):
            pdf.set_x(pdf.l_margin)
            pdf.set_font("Helvetica", font_style, font_size)
            pdf.cell(full_width(pdf), h, text, align='C', ln=1 if ln else 0)

        def center_multicell(text: str, h: float, font_style: str = '', font_size: float = 9):
            pdf.set_x(pdf.l_margin)
            pdf.set_font("Helvetica", font_style, font_size)
            pdf.multi_cell(full_width(pdf), h, text, align='C')

        # ---- Divider line (top of patient block)
        top_y = pdf.get_y()
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.2)
        pdf.line(pdf.l_margin, top_y, pdf.w - pdf.r_margin, top_y)
        pdf.ln(1.2 if not compact else 1.0)

        # ---- Patient & Invoice details (centered over full width)
        file_no = self.entries["Patient File No"].get().strip()
        pid = normalize_long_number(self.entries["ID"].get().strip())
        phone = self.entries["Phone"].get().strip()
        email = self.entries["Email"].get().strip()
        address = self.entries["Address"].get().strip()
        med_aid = self.entries["Medical Aid"].get().strip()
        med_plan = self.entries["Medical Aid Plan"].get().strip()
        member_no = self.entries["Membership No"].get().strip()

        def join_fields(pairs):
            parts = []
            for label, val in pairs:
                val = (val or "").strip()
                if val:
                    parts.append(f"{label}: {val}")
            return "  |  ".join(parts)

        name_line = ""
        if surname and name:
            name_line = f"{surname}, {name}"
        elif surname:
            name_line = surname
        elif name:
            name_line = name
        if file_no:
            name_line = f"{name_line}  (File: {file_no})" if name_line else f"(File: {file_no})"

        line_h = 5.0 if compact else 5.6
        if name_line:
            center_cell(name_line, line_h, 'B', 10 if not compact else 9)

        line2 = join_fields([("Invoice No", invoice_no), ("Invoice Date", invoice_date)])
        if line2:
            center_cell(line2, line_h, '', 9 if not compact else 8.2)

        line3 = join_fields([("ID", pid), ("Phone", phone), ("Email", email)])
        if line3:
            center_cell(line3, line_h, '', 9 if not compact else 8.2)

        if address:
            center_multicell(f"Address: {address}", line_h, '', 9 if not compact else 8.2)

        line5 = join_fields([("Medical Aid", med_aid), ("Plan", med_plan), ("Membership No", member_no)])
        if line5:
            center_cell(line5, line_h, '', 9 if not compact else 8.2)

        # ---- Divider line (bottom of patient block)
        pdf.ln(1.0 if not compact else 0.8)
        bottom_y = pdf.get_y()
        pdf.line(pdf.l_margin, bottom_y, pdf.w - pdf.r_margin, bottom_y)
        pdf.ln(1.2 if not compact else 1.0)

        # ---- Build rows from UI ----
        rows = []
        for line in self.lines:
            rows.append([
                widget_get_date_str(line['Date']),
                line['Service'].get(),
                line['ICD-10 Code'].get(),
                line['ICD-10 Secondary'].get(),
                line['Tariff'].get(),
                line['NAPPI'].get(),
                line['Qty'].get(),
                line['Price'].get(),
                line['Total'].get()
            ])

        # ===== Shared helpers for table =====
        def ensure_min_col_width(col_widths: List[float], target_idx: int, sample_texts: List[str],
                                 donors: List[int], min_keep_mm: float, font_size: float, pad_mm: float = 3.0) -> List[float]:
            pdf.set_font("Helvetica", '', font_size)
            need = max(pdf.get_string_width(t) for t in sample_texts) + pad_mm
            if col_widths[target_idx] >= need:
                return col_widths
            deficit = need - col_widths[target_idx]
            for d in donors:
                spare = max(col_widths[d] - min_keep_mm, 0)
                if spare <= 0:
                    continue
                take = min(spare, deficit)
                col_widths[d] -= take
                deficit -= take
                if deficit <= 0:
                    break
            col_widths[target_idx] = max(col_widths[target_idx], need)
            return col_widths

        def fitted_font_size_for_text(text: str, base_font: float, avail_w: float, pad: float = 1.6, min_font: float = 6.0) -> float:
            if not text:
                return base_font
            pdf.set_font("Helvetica", '', base_font)
            w = pdf.get_string_width(text)
            if w <= max(0.1, avail_w - pad):
                return base_font
            scale = (max(0.1, avail_w - pad)) / w
            return max(min_font, base_font * scale)

        def wrapped_height(text: str, col_width: float, line_hh: float, font_size: float) -> float:
            pad = 0.8
            if not text:
                return line_hh
            pdf.set_font("Helvetica", '', font_size)
            max_w = max(0.1, col_width - 1.6)
            s = str(text)
            words = s.split()
            lines = 1
            cur = ""

            def width(t: str) -> float:
                return pdf.get_string_width(t)

            if not words:
                cur = ""
                for ch in s:
                    t = cur + ch
                    if width(t) <= max_w:
                        cur = t
                    else:
                        lines += 1
                        cur = ch
                return lines * line_hh + pad

            for w in words:
                if width(w) > max_w:
                    for ch in w:
                        t = (cur + ch)
                        if width(t) <= max_w:
                            cur = t
                        else:
                            lines += 1
                            cur = ch
                else:
                    t = (cur + " " + w).strip()
                    if width(t) <= max_w:
                        cur = t
                    else:
                        lines += 1
                        cur = w
            return lines * line_hh + pad

        # ===== PORTRAIT =====
        if bool(self.portrait_var.get()):
            headers = ["Date", "Service", "ICD-10 Code", "ICD-10 Secondary", "Tariff", "NAPPI", "Qty", "Unit Price", "Total"]
            # wider ICD columns; NAPPI & Unit Price protected
            weights = [1.1, 1.8, 3.4, 3.4, 0.8, 1.2, 0.8, 1.2, 1.3]
            avail_w = pdf.w - pdf.l_margin - pdf.r_margin
            total_w = sum(weights)
            col_w = [avail_w * (w / total_w) for w in weights]

            WRAP_COLS = {1, 2, 3}
            RIGHT_ALIGN = {4, 6, 7, 8}
            NAPPI_COL = 5
            UNIT_PRICE_COL = 7
            TOTAL_COL = 8

            table_font = 8.2 if not self.compact_var.get() else 7.8
            header_h   = 5.6 if not self.compact_var.get() else 5.2
            row_h      = 4.4 if not self.compact_var.get() else 4.0
            try:
                min_font_portrait = float(self.min_portrait_font_var.get())
            except Exception:
                min_font_portrait = 6.2
            min_font_portrait = max(5.0, min(10.0, min_font_portrait))
            min_row_h  = 3.2
            min_cell_font = 6.0

            donors_portrait = [2, 3, 1, 4]
            col_w = ensure_min_col_width(col_w, NAPPI_COL, ["0000000000"], donors_portrait, 14.0, table_font)
            col_w = ensure_min_col_width(col_w, UNIT_PRICE_COL, ["R0000000.00", "0000000.00"], donors_portrait, 14.0, table_font)

            def table_total_height(fsize: float, base_h: float, widths: List[float]) -> float:
                h = header_h
                for r in rows:
                    heights = []
                    for i, txt in enumerate(r):
                        if i in WRAP_COLS:
                            heights.append(wrapped_height(str(txt or ""), widths[i], base_h, fsize) + 0.6)
                        else:
                            heights.append(base_h)
                    h += max(heights) if heights else base_h
                return h + header_h

            def free_height() -> float:
                return (pdf.h - pdf.b_margin) - pdf.get_y()

            while table_total_height(table_font, row_h, col_w) > free_height():
                shrunk = False
                if table_font > min_font_portrait:
                    table_font -= 0.2
                    shrunk = True
                if row_h > min_row_h:
                    row_h -= 0.08
                    shrunk = True
                if not shrunk:
                    cur_y = pdf.get_y()
                    if pdf.l_margin > 6 or pdf.r_margin > 6:
                        pdf.set_margins(max(pdf.l_margin - 1, 6), pdf.t_margin, max(pdf.r_margin - 1, 6))
                        avail_w = pdf.w - pdf.l_margin - pdf.r_margin
                        col_w = [avail_w * (w / total_w) for w in weights]
                        col_w = ensure_min_col_width(col_w, NAPPI_COL, ["0000000000"], donors_portrait, 14.0, table_font)
                        col_w = ensure_min_col_width(col_w, UNIT_PRICE_COL, ["R0000000.00", "0000000.00"], donors_portrait, 14.0, table_font)
                        pdf.set_y(cur_y)
                    else:
                        break
                col_w = ensure_min_col_width(col_w, NAPPI_COL, ["0000000000"], donors_portrait, 14.0, table_font)
                col_w = ensure_min_col_width(col_w, UNIT_PRICE_COL, ["R0000000.00", "0000000.00"], donors_portrait, 14.0, table_font)

            # draw header
            pdf.set_font("Helvetica", 'B', table_font)
            for w, htxt in zip(col_w, headers):
                pdf.cell(w, header_h, htxt, 1, align='C')
            pdf.ln(header_h)

            # x anchors
            x_pos = [pdf.l_margin]
            for w in col_w[:-1]:
                x_pos.append(x_pos[-1] + w)

            pdf.set_font("Helvetica", '', table_font)

            def draw_row(vals: List[str]):
                y0 = pdf.get_y()
                # conservative heights with safety bump to prevent spillover
                heights = []
                for i, text in enumerate(vals):
                    if i in WRAP_COLS:
                        h_est = wrapped_height(str(text or ""), col_w[i], row_h, table_font) + 0.6
                        heights.append(h_est)
                    else:
                        heights.append(row_h)
                rh = max(heights) if heights else row_h

                for i, text in enumerate(vals):
                    x = x_pos[i]; w = col_w[i]
                    pdf.rect(x, y0, w, rh)
                    txt = str(text or "")
                    if i in WRAP_COLS:
                        pdf.set_xy(x + 0.8, y0 + 0.3)
                        pdf.multi_cell(w - 1.6, row_h, txt, border=0, align='L')
                        pdf.set_xy(x_pos[0], y0)
                    else:
                        base_align = 'R' if i in {UNIT_PRICE_COL, 8, 4, 6} else ('C' if i == 5 else 'C')
                        fit_font = fitted_font_size_for_text(txt, table_font, w, pad=1.6, min_font=min_cell_font)
                        pdf.set_xy(x, y0)
                        pdf.set_font("Helvetica", '', fit_font)
                        pdf.cell(w, rh, txt, border=0, align=base_align)
                        pdf.set_font("Helvetica", '', table_font)
                pdf.set_xy(x_pos[0], y0 + rh)

            # totals calc (VAT-inclusive input)
            gross = Decimal('0.00')
            for r in rows:
                draw_row(r)
                try:
                    gross += money_to_decimal(r[-1])
                except InvalidOperation:
                    pass

            ex = (gross / (Decimal(1) + self.vat_rate)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP) if self.vat_rate else gross
            vat = (gross - ex).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP) if self.vat_rate else Decimal('0.00')

            pdf.ln(0.6)
            pdf.set_font("Helvetica", 'B', max(table_font - 0.1, 6.0))
            totals_line = f"Excl. VAT: {fmt_money(ex)}   |   VAT @{int(self.vat_rate*100)}%: {fmt_money(vat)}   |   Total (Incl. VAT): {fmt_money(gross)}"
            pdf.cell(0, header_h, totals_line, ln=True, align='R')

        # ===== LANDSCAPE =====
        else:
            headers = ["Date", "Service", "ICD-10 Code", "ICD-10 Secondary", "Tariff", "NAPPI", "Qty", "Unit Price", "Total"]
            weights = [0.9, 1.8, 3.8, 3.8, 0.7, 1.3, 0.6, 1.1, 1.0]
            avail_w = pdf.w - pdf.l_margin - pdf.r_margin
            col_widths = [avail_w * (w / sum(weights)) for w in weights]
            WRAP_COLS = {1, 2, 3}
            table_font = 8.6 if not self.compact_var.get() else 8.0
            base_row_h = 5.4 if not self.compact_var.get() else 4.8
            min_cell_font = 6.0

            donors_land = [2, 3, 1, 4]
            NAPPI_COL = 5
            UNIT_PRICE_COL = 7
            col_widths = ensure_min_col_width(col_widths, NAPPI_COL, ["0000000000"], donors_land, 14.0, table_font)
            col_widths = ensure_min_col_width(col_widths, UNIT_PRICE_COL, ["R0000000.00", "0000000.00"], donors_land, 14.0, table_font)

            pdf.set_font("Helvetica", 'B', table_font)
            for w, h in zip(col_widths, headers):
                pdf.cell(w, 7 if not self.compact_var.get() else 6, h, 1, align='C')
            pdf.ln()

            pdf.set_font("Helvetica", '', table_font)
            x_positions = [pdf.l_margin]
            for w in col_widths[:-1]:
                x_positions.append(x_positions[-1] + w)

            def draw_row(row_vals: List[str]):
                y0 = pdf.get_y()
                heights = []
                for i, text in enumerate(row_vals):
                    if i in WRAP_COLS:
                        h_est = wrapped_height(str(text or ""), col_widths[i], base_row_h, table_font) + 0.6
                        heights.append(h_est)
                    else:
                        heights.append(base_row_h)
                total_h = max(heights) if heights else base_row_h

                for i in range(len(col_widths)):
                    x = x_positions[i]; w = col_widths[i]
                    pdf.rect(x, y0, w, total_h)
                    text = str(row_vals[i] or "")
                    if i in WRAP_COLS:
                        pdf.set_xy(x + 0.8, y0 + 0.5)
                        pdf.multi_cell(w - 1.6, base_row_h, text, border=0, align='L')
                        pdf.set_xy(x_positions[0], y0)
                    else:
                        base_align = 'R' if i in {UNIT_PRICE_COL, 8, 4, 6} else ('C' if i == 5 else 'C')
                        fit_font = fitted_font_size_for_text(text, table_font, w, pad=1.6, min_font=min_cell_font)
                        pdf.set_xy(x, y0)
                        pdf.set_font("Helvetica", '', fit_font)
                        pdf.cell(w, total_h, text, border=0, align=base_align)
                        pdf.set_font("Helvetica", '', table_font)
                pdf.set_xy(x_positions[0], y0 + total_h)

            gross = Decimal('0.00')
            for r in rows:
                draw_row(r)
                try:
                    gross += money_to_decimal(r[-1])
                except InvalidOperation:
                    pass

            if self.vat_rate == 0:
                ex = gross
                vat = Decimal('0.00')
            else:
                ex = (gross / (Decimal(1) + self.vat_rate)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                vat = (gross - ex).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

            pdf.ln(1 if not self.compact_var.get() else 0.5)
            pdf.set_font("Helvetica", 'B', table_font + (0.7 if not self.compact_var.get() else 0.5))
            totals_line = f"Excl. VAT: {fmt_money(ex)}   |   VAT @{int(self.vat_rate*100)}%: {fmt_money(vat)}   |   Total (Incl. VAT): {fmt_money(gross)}"
            pdf.cell(0, 7 if not self.compact_var.get() else 6, totals_line, ln=True, align='R')

        # Save as <INV>-<Surname>-<Name>-<Date>.pdf
        clean_surname = self.entries["Surname"].get().strip().replace(" ", "_")
        clean_name = self.entries["Name"].get().strip().replace(" ", "_")
        default_name = f"{invoice_no}_{clean_surname}_{clean_name}_{invoice_date}.pdf"
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default_name, filetypes=[("PDF", "*.pdf")])
        if filepath:
            try:
                pdf.output(filepath)
                self.last_pdf_path = filepath
                persist_invoice_if_new(invoice_no)   # persist counter now
                self._assign_peek_invoice_number()   # show next
                messagebox.showinfo("Saved", f"Invoice saved to:\n{filepath}")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

    def print_pdf(self):
        """
        Send the last generated PDF straight to a physical printer.
        - Windows: prefers pywin32 'printto', falls back to os.startfile('print').
        - macOS/Linux: uses 'lp' or 'lpr' to the default printer (no preview).
        """
        if not self.last_pdf_path or not os.path.exists(self.last_pdf_path):
            messagebox.showerror(
                "No PDF to print",
                "Please generate the invoice first (Generate PDF), then click Print."
            )
            return

        path = self.last_pdf_path
        try:
            system = platform.system()

            if system == "Windows":
                try:
                    import win32print  # type: ignore
                    import win32api    # type: ignore
                    printer_name = win32print.GetDefaultPrinter()
                    subprocess.Popen(['cmd', '/c', 'start', '', '/MIN', path], shell=False)
                    # Alternative (may require default app association):
                    # win32api.ShellExecute(0, "printto", path, f'"{printer_name}"', ".", 0)
                except Exception:
                    os.startfile(path, "print")

            elif system == "Darwin":
                import shutil
                if shutil.which("lp"):
                    subprocess.run(["lp", path], check=False)
                else:
                    subprocess.run(["lpr", path], check=False)

            else:
                import shutil
                if shutil.which("lp"):
                    subprocess.run(["lp", path], check=False)
                else:
                    subprocess.run(["lpr", path], check=False)

            messagebox.showinfo("Printing", "Invoice sent to your default printer.")
        except Exception as e:
            messagebox.showerror("Print Error", f"Couldn't print the file.\n\n{e}")

    # ----- Drafts -----

    def save_draft(self):
        draft = {
            "vat_rate": float(self.vat_rate),
            "min_portrait_font": float(self.min_portrait_font_var.get()),
            "invoice_no": self.entries["Invoice No"].get(),
            "invoice_date": self._get_invoice_date(),
            "patient": {k: self.entries[k].get() for k in PATIENT_FIELDS},
            "lines": [{
                k: (widget_get_date_str(v) if k == "Date" else v.get())
                for k, v in line.items() if not k.startswith("_")
            } for line in self.lines]
        }
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")], initialfile="invoice_draft.json")
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(draft, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("Saved", "Draft saved successfully.")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

    def load_draft(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not path or not os.path.exists(path):
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return

        self.ui_ready = False

        if "vat_rate" in data:
            try:
                self.vat_rate = Decimal(str(data["vat_rate"]))
                self.vat_var.set(f"{(self.vat_rate*100):.0f}")
            except InvalidOperation:
                self.vat_rate = DEFAULT_VAT_RATE
                self.vat_var.set(f"{(DEFAULT_VAT_RATE*100):.0f}")

        if "min_portrait_font" in data:
            try:
                self.min_portrait_font_var.set(float(data["min_portrait_font"]))
            except Exception:
                self.min_portrait_font_var.set(6.2)

        inv_no = data.get("invoice_no")
        if not inv_no:
            inv_no = format_invoice_no(peek_next_invoice_number())
        self.entries["Invoice No"].configure(state="normal")
        self.entries["Invoice No"].delete(0, tk.END)
        self.entries["Invoice No"].insert(0, inv_no)
        self.entries["Invoice No"].configure(state="readonly")

        inv_date = data.get("invoice_date", datetime.today().strftime(DATE_FMT))
        widget_set_date(self.entries["Invoice Date"], inv_date)

        for k, v in data.get('patient', {}).items():
            if k in self.entries:
                self.entries[k].delete(0, tk.END)
                self.entries[k].insert(0, v)

        for widgets in self.lines:
            for widget in widgets.values():
                try:
                    widget.destroy()
                except Exception:
                    pass
        self.lines.clear()

        for line_data in data.get('lines', []):
            self.add_line_item(preset=line_data)

        self.ui_ready = True
        self.recalc_totals()

    def reset_form(self):
        self.ui_ready = False
        self._assign_peek_invoice_number()
        widget_set_date(self.entries["Invoice Date"], datetime.today().strftime(DATE_FMT))

        for key in PATIENT_FIELDS:
            self.entries[key].delete(0, tk.END)

        self.vat_rate = DEFAULT_VAT_RATE
        self.vat_var.set(f"{(DEFAULT_VAT_RATE*100):.0f}")
        self.min_portrait_font_var.set(6.2)

        for widgets in self.lines:
            for widget in widgets.values():
                try:
                    widget.destroy()
                except Exception:
                    pass
        self.lines.clear()

        self.add_line_item()
        self.ui_ready = True
        self.recalc_totals()

# =========================
# Main
# =========================

if __name__ == '__main__':
    root = tk.Tk()
    try:
        root.call("tk", "scaling", 1.25)
    except Exception:
        pass
    app = InvoiceApp(root)
    root.mainloop()















































