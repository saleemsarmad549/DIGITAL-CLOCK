"""
Data Tools Suite
=================
A single merged application that combines two previously-separate GUIs into
one window using Chrome-style tabs (ttk.Notebook):

  Tab 1: "DTA -> Excel"   - convert .DTA files into clean .xlsx workbooks
  Tab 2: "Import to SQL"  - import Excel files into a MySQL database

Both tabs keep 100% of their original functionality; they've just been
converted from standalone tk.Tk windows into tk.Frame "pages" that live
inside one shared Notebook, with a single shared theme/style setup.
"""

import os
import re
import hashlib
import shutil
import locale
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from sqlalchemy import create_engine, text

try:
    import gamry_parser
    _HAS_GAMRY_PARSER = True
except ImportError:
    _HAS_GAMRY_PARSER = False


def _robust_atof(value: str) -> float:
    """Drop-in replacement for locale.atof that also handles European-style
    decimal-comma numbers (e.g. '2,00000E+003' meaning 2000.0), which some
    Gamry instruments/software export regardless of the running machine's
    locale settings. Falls back to the real locale.atof for anything else."""
    s = value.strip()
    try:
        return float(s)
    except ValueError:
        pass
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return locale.atof(s)


if _HAS_GAMRY_PARSER:
    locale.atof = _robust_atof

# ---------------------------------------------------------------------------
# SHARED THEME
# ---------------------------------------------------------------------------

BG = "#0f1720"          # app background (deep navy)
PANEL = "#17212b"       # card / panel background
PANEL_ALT = "#1e2b38"   # slightly lighter panel
ACCENT = "#33c9a3"      # teal accent
ACCENT_DARK = "#249c80"
DANGER = "#e05a5a"
DANGER_DARK = "#b84545"
TEXT = "#e7edf2"
TEXT_MUTED = "#8fa3b3"
FONT = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 11, "bold")
FONT_TITLE = ("Segoe UI", 16, "bold")
FONT_SMALL = ("Segoe UI", 9)

MAX_SHEET_NAME_LEN = 31  # Excel hard limit
SYSTEM_DBS = {"information_schema", "mysql", "performance_schema", "sys"}

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CELL_DATA_ROOT = os.path.join(APP_DIR, "cell_data")


def build_shared_style(root: tk.Tk):
    """One ttk.Style setup shared by both tabs (styles are global per
    application, so this only needs to run once from the top-level window)."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure("TFrame", background=BG)
    style.configure("Panel.TFrame", background=PANEL)
    style.configure("TLabel", background=BG, foreground=TEXT, font=FONT)
    style.configure("Panel.TLabel", background=PANEL, foreground=TEXT, font=FONT)
    style.configure("Muted.TLabel", background=PANEL, foreground=TEXT_MUTED, font=FONT_SMALL)
    style.configure("Title.TLabel", background=BG, foreground=TEXT, font=FONT_TITLE)

    style.configure("TEntry", fieldbackground=PANEL_ALT, foreground=TEXT,
                     insertcolor=TEXT, borderwidth=0)
    style.map("TEntry", fieldbackground=[("focus", PANEL_ALT)])

    style.configure("Accent.TButton", background=ACCENT, foreground="#06231c",
                     font=FONT_BOLD, borderwidth=0, padding=10)
    style.map("Accent.TButton", background=[("active", ACCENT_DARK)])

    style.configure("Ghost.TButton", background=PANEL_ALT, foreground=TEXT,
                     font=FONT, borderwidth=0, padding=8)
    style.map("Ghost.TButton", background=[("active", "#26374a")])

    style.configure("Danger.TButton", background=DANGER, foreground="#2a0d0d",
                     font=FONT, borderwidth=0, padding=8)
    style.map("Danger.TButton", background=[("active", DANGER_DARK)])

    style.configure("TRadiobutton", background=PANEL, foreground=TEXT, font=FONT)
    style.map("TRadiobutton", background=[("active", PANEL)])

    style.configure("TProgressbar", background=ACCENT, troughcolor=PANEL_ALT,
                     borderwidth=0)

    style.configure("Treeview", background=PANEL_ALT, fieldbackground=PANEL_ALT,
                     foreground=TEXT, rowheight=28, borderwidth=0, font=FONT)
    style.configure("Treeview.Heading", background=PANEL, foreground=TEXT_MUTED,
                     font=FONT_SMALL, borderwidth=0)
    style.map("Treeview", background=[("selected", ACCENT_DARK)],
               foreground=[("selected", "#06231c")])

    # Chrome-style tab strip
    style.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(8, 8, 8, 0))
    style.configure("TNotebook.Tab", background=PANEL, foreground=TEXT_MUTED,
                     font=FONT_BOLD, padding=(18, 10), borderwidth=0)
    style.map("TNotebook.Tab",
              background=[("selected", BG)],
              foreground=[("selected", TEXT)])
    return style


# ---------------------------------------------------------------------------
# DTA -> EXCEL: PARSING / CONVERSION HELPERS
# ---------------------------------------------------------------------------


def sanitize_sheet_name(name: str, used: set) -> str:
    """Excel sheet names: <=31 chars, no []:*?/\\ characters, must be unique."""
    name = re.sub(r'[\[\]:\*\?/\\]', "_", name).strip() or "Sheet"
    name = name[:MAX_SHEET_NAME_LEN]
    base = name
    i = 1
    while name in used:
        suffix = f"_{i}"
        name = base[: MAX_SHEET_NAME_LEN - len(suffix)] + suffix
        i += 1
    used.add(name)
    return name


def _try_gamry_parser(filepath: str):
    """Use the dedicated 'gamry-parser' library (pip install gamry-parser) to
    properly parse a genuine Gamry EXPLAIN-format .DTA file."""
    if not _HAS_GAMRY_PARSER:
        return None
    try:
        gp = gamry_parser.GamryParser()
        gp.load(filename=filepath)
    except Exception:
        return None

    sheets = {}

    header = gp.get_header()
    if header:
        rows = []
        for key, value in header.items():
            if isinstance(value, dict):
                value = ", ".join(f"{k}={v}" for k, v in value.items())
            rows.append({"Field": key, "Value": value})
        sheets["Header"] = pd.DataFrame(rows)

    ocv = gp.get_ocv_curve()
    if ocv is not None and not ocv.empty:
        sheets["OCV_Curve"] = ocv.reset_index()

    curves = gp.get_curves()
    if curves:
        if len(curves) == 1:
            sheets["Curve"] = curves[0].reset_index()
        else:
            for i, df in enumerate(curves, start=1):
                sheets[f"Curve_{i}"] = df.reset_index()

    return sheets if sheets else None


def _try_read_stata(filepath: str):
    """Some tools save statistical data with a .dta extension (Stata format)."""
    try:
        df = pd.read_stata(filepath)
        if not df.empty:
            return {"Data": df}
    except Exception:
        pass
    return None


def _read_lines(filepath: str):
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            with open(filepath, "r", encoding=enc, errors="strict") as f:
                return f.read().splitlines()
        except Exception:
            continue
    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        return f.read().splitlines()


def _try_read_gamry_tables(filepath: str):
    """Instrument-exported .DTA files (tab-delimited, TABLE-section format)."""
    lines = _read_lines(filepath)
    tables = {}
    i = 0
    n = len(lines)
    found_any = False

    while i < n:
        parts = lines[i].split("\t")
        if len(parts) >= 2 and parts[1].strip().upper() == "TABLE":
            found_any = True
            table_name = parts[0].strip() or f"TABLE_{len(tables) + 1}"
            i += 1
            if i >= n:
                break
            header = [h.strip() for h in lines[i].split("\t")]
            i += 1
            if i < n:
                probe = lines[i].split("\t")
                looks_numeric = _row_is_mostly_numeric(probe)
                if not looks_numeric:
                    i += 1  # units row, skip

            rows = []
            while i < n and lines[i].strip() != "":
                row_parts = lines[i].split("\t")
                probe = row_parts[1].strip().upper() if len(row_parts) > 1 else ""
                if probe == "TABLE":
                    break
                rows.append(row_parts)
                i += 1

            if rows:
                width = len(header)
                normalized = [
                    (r + [""] * width)[:width] if len(r) < width else r[:width]
                    for r in rows
                ]
                df = pd.DataFrame(normalized, columns=header)
                if df.columns[0] == "" and df.iloc[:, 0].eq("").all():
                    df = df.iloc[:, 1:]
                df = df.apply(pd.to_numeric, errors="ignore")
                tables[table_name] = df
        else:
            i += 1

    return tables if found_any and tables else None


def _row_is_mostly_numeric(parts) -> bool:
    vals = [p for p in parts if p.strip() != ""]
    if not vals:
        return False
    numeric = 0
    for v in vals:
        try:
            float(v)
            numeric += 1
        except ValueError:
            pass
    return numeric >= max(1, len(vals) // 2)


def _try_generic_delimited(filepath: str):
    """Last-resort fallback: find the most plausible header row and read
    everything below it as a table."""
    lines = _read_lines(filepath)
    if not lines:
        return None

    delimiter = "\t" if any("\t" in ln for ln in lines[:20]) else ","

    counts = [len(ln.split(delimiter)) for ln in lines]
    header_idx = None
    for idx in range(len(lines) - 3):
        c = counts[idx]
        if c < 2:
            continue
        if counts[idx + 1] == c and counts[idx + 2] == c:
            header_idx = idx
            break
    if header_idx is None:
        return None

    header = [h.strip() or f"col_{j}" for j, h in enumerate(lines[header_idx].split(delimiter))]
    data_lines = [ln for ln in lines[header_idx + 1:] if ln.strip() != ""]
    rows = [ln.split(delimiter) for ln in data_lines]
    width = len(header)
    normalized = [
        (r + [""] * width)[:width] if len(r) < width else r[:width]
        for r in rows
    ]
    if not normalized:
        return None
    df = pd.DataFrame(normalized, columns=header)
    df = df.apply(pd.to_numeric, errors="ignore")
    return {"Data": df}


def convert_dta_to_dataframes(filepath: str):
    """Try, in order: gamry_parser, Stata .dta, hand-rolled Gamry TABLE
    parser, then generic delimited text."""
    result = _try_gamry_parser(filepath)
    if result:
        return result

    result = _try_read_stata(filepath)
    if result:
        return result

    result = _try_read_gamry_tables(filepath)
    if result:
        return result

    result = _try_generic_delimited(filepath)
    if result:
        return result

    hint = "" if _HAS_GAMRY_PARSER else (
        " (Note: the 'gamry-parser' library is not installed - run "
        "'pip install gamry-parser' for better support of real Gamry files.)"
    )
    raise ValueError("Could not recognize the file's structure (not a valid "
                      "Gamry EXPLAIN file, not Stata, and no consistent "
                      "delimited table was found)." + hint)


def write_dataframes_to_excel(dataframes: dict, output_path: str):
    used = set()
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in dataframes.items():
            sheet = sanitize_sheet_name(name, used)
            df.to_excel(writer, sheet_name=sheet, index=False)


# ---------------------------------------------------------------------------
# IMPORT TO SQL: HELPERS
# ---------------------------------------------------------------------------


def sanitize_name(name: str) -> str:
    name = name.strip().upper()
    name = re.sub(r"[^A-Z0-9]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    if not name:
        name = "UNNAMED"
    if name[0].isdigit():
        name = f"T_{name}"
    return name


def sanitize_columns(columns) -> list:
    seen = {}
    clean = []
    for col in columns:
        base = sanitize_name(str(col))
        if base in seen:
            seen[base] += 1
            base = f"{base}_{seen[base]}"
        else:
            seen[base] = 0
        clean.append(base)
    return clean


MYSQL_MAX_IDENTIFIER_LEN = 64


def build_table_name(*parts) -> str:
    """Join sanitized parts into a MySQL-safe table name, truncating with a
    content hash if the joined name would exceed 64 characters."""
    clean_parts = [sanitize_name(p) for p in parts if p]
    full = "__".join(clean_parts)
    if len(full) <= MYSQL_MAX_IDENTIFIER_LEN:
        return full
    digest = hashlib.md5(full.encode("utf-8")).hexdigest()[:8]
    keep = MYSQL_MAX_IDENTIFIER_LEN - len(digest) - 1
    return f"{full[:keep].rstrip('_')}_{digest}"


def import_excel_file(engine, filepath: str, log_fn, table_prefix: str = "",
                       replace_existing: bool = True):
    """Import every sheet of one Excel file into its own table."""
    file_stub = sanitize_name(os.path.splitext(os.path.basename(filepath))[0])
    try:
        sheets = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
    except Exception as exc:
        log_fn(f"[ERROR] Could not read {os.path.basename(filepath)}: {exc}")
        return 0

    tables_written = 0
    for sheet_name, df in sheets.items():
        if df.empty:
            log_fn(f"[SKIP] Empty sheet '{sheet_name}' in {os.path.basename(filepath)}")
            continue

        df.columns = sanitize_columns(df.columns)
        sheet_part = sanitize_name(sheet_name) if len(sheets) > 1 else None
        table_name = build_table_name(table_prefix, file_stub, sheet_part)
        if_exists = "replace" if replace_existing else "append"

        try:
            df.to_sql(table_name, engine, if_exists=if_exists, index=False,
                      chunksize=1000, method="multi")
            log_fn(f"[OK] {os.path.basename(filepath)} -> '{table_name}' "
                   f"({len(df)} rows, {len(df.columns)} cols)")
            tables_written += 1
        except Exception as exc:
            log_fn(f"[ERROR] Failed writing table '{table_name}': {exc}")

    return tables_written


def drop_tables_with_prefix(engine, prefix: str, log_fn=None) -> int:
    """Drop every table whose name starts with the given prefix."""
    if engine is None or not prefix:
        return 0
    dropped = 0
    try:
        with engine.connect() as c:
            result = c.execute(text("SHOW TABLES"))
            all_tables = [row[0] for row in result]
            matches = [t for t in all_tables if t.upper().startswith(prefix.upper())]
            for t in matches:
                c.execute(text(f"DROP TABLE IF EXISTS `{t}`"))
                dropped += 1
            c.commit()
    except Exception as exc:
        if log_fn:
            log_fn(f"[ERROR] Could not drop tables for prefix '{prefix}': {exc}")
    return dropped


def fs_safe_name(name: str) -> str:
    name = (name or "").strip()
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name or "unnamed"


def db_dir(db_name: str) -> str:
    return os.path.join(CELL_DATA_ROOT, fs_safe_name(db_name))


def cell_dir(db_name: str, cell_name: str) -> str:
    return os.path.join(db_dir(db_name), fs_safe_name(cell_name))


def list_cells(db_name: str) -> list:
    d = db_dir(db_name)
    if not os.path.isdir(d):
        return []
    return sorted(c for c in os.listdir(d) if os.path.isdir(os.path.join(d, c)))


def list_cell_files(db_name: str, cell_name: str) -> list:
    d = cell_dir(db_name, cell_name)
    if not os.path.isdir(d):
        return []
    return sorted(f for f in os.listdir(d) if f.lower().endswith((".xlsx", ".xls")))


def create_cell(db_name: str, cell_name: str):
    os.makedirs(cell_dir(db_name, cell_name), exist_ok=True)


def delete_cell(db_name: str, cell_name: str):
    shutil.rmtree(cell_dir(db_name, cell_name), ignore_errors=True)


def copy_files_into_cell(db_name: str, cell_name: str, filepaths) -> list:
    dest_dir = cell_dir(db_name, cell_name)
    os.makedirs(dest_dir, exist_ok=True)
    copied = []
    for fp in filepaths:
        dest = os.path.join(dest_dir, os.path.basename(fp))
        try:
            shutil.copy2(fp, dest)
            copied.append(dest)
        except Exception:
            pass
    return copied


def delete_cell_file(db_name: str, cell_name: str, filename: str):
    try:
        os.remove(os.path.join(cell_dir(db_name, cell_name), filename))
    except OSError:
        pass


def delete_db_folder(db_name: str):
    shutil.rmtree(db_dir(db_name), ignore_errors=True)


# ---------------------------------------------------------------------------
# TAB 1: DTA -> EXCEL CONVERTER
# ---------------------------------------------------------------------------


class DtaToExcelTab(tk.Frame):
    """Import .DTA files, export clean .xlsx workbooks."""

    def __init__(self, parent):
        super().__init__(parent, bg=BG)

        self.files = []
        self.export_mode = tk.StringVar(value="separate")
        self.export_folder = tk.StringVar(value="")
        self.export_combined_path = tk.StringVar(value="")

        self._build_header()
        self._build_body()

    def _build_header(self):
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=24, pady=(20, 10))
        tk.Label(header, text="\U0001F501 DTA \u2192 Excel Converter", bg=BG, fg=TEXT,
                 font=FONT_TITLE).pack(side="left")
        tk.Label(header, text="Import .DTA files, export clean .xlsx workbooks",
                 bg=BG, fg=TEXT_MUTED, font=FONT_SMALL).pack(side="left", padx=(16, 0))

    def _build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=(0, 24))

        files_panel = tk.Frame(body, bg=PANEL, padx=16, pady=16)
        files_panel.pack(fill="x", pady=(0, 16))

        ttk.Label(files_panel, text="Step 1 \u2014 Import .DTA files", style="Panel.TLabel",
                  font=FONT_BOLD).pack(anchor="w", pady=(0, 8))
        ttk.Label(files_panel, text="Add as many files as you like. They'll all be converted together.",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 8))

        list_frame = tk.Frame(files_panel, bg=PANEL)
        list_frame.pack(fill="both", expand=True)

        self.files_list = tk.Listbox(list_frame, bg=PANEL_ALT, fg=TEXT, height=8,
                                      selectmode=tk.EXTENDED,
                                      selectbackground=ACCENT_DARK, selectforeground="#06231c",
                                      borderwidth=0, highlightthickness=0, font=FONT_SMALL,
                                      activestyle="none")
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.files_list.yview)
        self.files_list.configure(yscrollcommand=list_scroll.set)
        self.files_list.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y")

        self.no_files_label = ttk.Label(files_panel, text="No files added yet.",
                                         style="Muted.TLabel")
        self.no_files_label.pack(anchor="w", pady=(6, 0))

        files_btn_row = tk.Frame(files_panel, bg=PANEL)
        files_btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(files_btn_row, text="+ Import file(s)\u2026", style="Accent.TButton",
                   command=self.import_files).pack(side="left")
        ttk.Button(files_btn_row, text="Remove selected", style="Ghost.TButton",
                   command=self.remove_selected).pack(side="left", padx=8)
        ttk.Button(files_btn_row, text="Clear all", style="Danger.TButton",
                   command=self.clear_files).pack(side="left")

        export_panel = tk.Frame(body, bg=PANEL, padx=16, pady=16)
        export_panel.pack(fill="x", pady=(0, 16))

        ttk.Label(export_panel, text="Step 2 \u2014 Export to", style="Panel.TLabel",
                  font=FONT_BOLD).pack(anchor="w", pady=(0, 8))

        mode_row = tk.Frame(export_panel, bg=PANEL)
        mode_row.pack(fill="x", pady=(0, 10))
        ttk.Radiobutton(mode_row, text="One .xlsx file per import (into a folder)",
                         variable=self.export_mode, value="separate",
                         command=self._refresh_export_controls).pack(anchor="w")
        ttk.Radiobutton(mode_row, text="Combine everything into a single workbook (one sheet per file/table)",
                         variable=self.export_mode, value="combined",
                         command=self._refresh_export_controls).pack(anchor="w", pady=(4, 0))

        self.dest_row = tk.Frame(export_panel, bg=PANEL)
        self.dest_row.pack(fill="x")

        self.dest_entry = ttk.Entry(self.dest_row, textvariable=self.export_folder, width=52)
        self.dest_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(self.dest_row, text="Browse\u2026", style="Ghost.TButton",
                   command=self.choose_destination).pack(side="left", padx=(8, 0))

        convert_panel = tk.Frame(body, bg=PANEL, padx=16, pady=16)
        convert_panel.pack(fill="both", expand=True)

        ttk.Label(convert_panel, text="Step 3 \u2014 Convert", style="Panel.TLabel",
                  font=FONT_BOLD).pack(anchor="w", pady=(0, 8))

        self.convert_btn = ttk.Button(convert_panel, text="Convert to Excel",
                                       style="Accent.TButton", command=self.start_convert)
        self.convert_btn.pack(anchor="e")

        self.progress = ttk.Progressbar(convert_panel, mode="indeterminate")

        self.log_box = tk.Text(convert_panel, height=10, bg="#0c141c", fg=TEXT_MUTED,
                                font=("Consolas", 9), borderwidth=0, highlightthickness=0)
        self.log_box.pack(fill="both", expand=True, pady=(10, 0))
        self.log_box.configure(state="disabled")

        self._refresh_files_list()
        self._refresh_export_controls()

    def import_files(self):
        paths = filedialog.askopenfilenames(
            title="Select .DTA files to import",
            filetypes=[("DTA files", "*.dta *.DTA"), ("All files", "*.*")]
        )
        if not paths:
            return
        for p in paths:
            if p not in self.files:
                self.files.append(p)
        self._refresh_files_list()

    def remove_selected(self):
        sel = list(self.files_list.curselection())
        if not sel:
            return
        for idx in reversed(sel):
            del self.files[idx]
        self._refresh_files_list()

    def clear_files(self):
        self.files.clear()
        self._refresh_files_list()

    def _refresh_files_list(self):
        self.files_list.delete(0, tk.END)
        for p in self.files:
            self.files_list.insert(tk.END, os.path.basename(p))
        if self.files:
            self.no_files_label.pack_forget()
        else:
            self.no_files_label.pack(anchor="w", pady=(6, 0))

    def _refresh_export_controls(self):
        if self.export_mode.get() == "separate":
            self.dest_entry.configure(textvariable=self.export_folder)
        else:
            self.dest_entry.configure(textvariable=self.export_combined_path)

    def choose_destination(self):
        if self.export_mode.get() == "separate":
            folder = filedialog.askdirectory(title="Choose a folder to export .xlsx files into")
            if folder:
                self.export_folder.set(folder)
        else:
            path = filedialog.asksaveasfilename(
                title="Save combined workbook as",
                defaultextension=".xlsx",
                filetypes=[("Excel workbook", "*.xlsx")],
                initialfile="combined_export.xlsx",
            )
            if path:
                self.export_combined_path.set(path)

    def _log(self, msg):
        def append():
            self.log_box.configure(state="normal")
            self.log_box.insert(tk.END, msg + "\n")
            self.log_box.see(tk.END)
            self.log_box.configure(state="disabled")
        self.after(0, append)

    def start_convert(self):
        if not self.files:
            messagebox.showwarning("No files", "Import at least one .DTA file first.")
            return

        mode = self.export_mode.get()
        if mode == "separate" and not self.export_folder.get().strip():
            messagebox.showwarning("No destination", "Choose a folder to export the files into.")
            return
        if mode == "combined" and not self.export_combined_path.get().strip():
            messagebox.showwarning("No destination", "Choose where to save the combined workbook.")
            return

        self.convert_btn.state(["disabled"])
        self.progress.pack(fill="x", pady=(4, 8))
        self.progress.start(10)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", tk.END)
        self.log_box.configure(state="disabled")

        thread = threading.Thread(target=self._run_convert, args=(mode,), daemon=True)
        thread.start()

    def _run_convert(self, mode):
        ok_count = 0
        fail_count = 0

        if mode == "separate":
            out_dir = self.export_folder.get().strip()
            os.makedirs(out_dir, exist_ok=True)
            for filepath in self.files:
                stub = os.path.splitext(os.path.basename(filepath))[0]
                out_path = os.path.join(out_dir, f"{stub}.xlsx")
                try:
                    dataframes = convert_dta_to_dataframes(filepath)
                    write_dataframes_to_excel(dataframes, out_path)
                    self._log(f"[OK] {os.path.basename(filepath)} \u2192 {out_path} "
                               f"({len(dataframes)} sheet(s))")
                    ok_count += 1
                except Exception as exc:
                    self._log(f"[ERROR] {os.path.basename(filepath)}: {exc}")
                    fail_count += 1
        else:
            combined = {}
            used_names = set()
            for filepath in self.files:
                stub = os.path.splitext(os.path.basename(filepath))[0]
                try:
                    dataframes = convert_dta_to_dataframes(filepath)
                    for table_name, df in dataframes.items():
                        label = stub if len(dataframes) == 1 else f"{stub}_{table_name}"
                        sheet = sanitize_sheet_name(label, used_names)
                        combined[sheet] = df
                    self._log(f"[OK] {os.path.basename(filepath)} \u2192 "
                               f"{len(dataframes)} sheet(s) queued")
                    ok_count += 1
                except Exception as exc:
                    self._log(f"[ERROR] {os.path.basename(filepath)}: {exc}")
                    fail_count += 1

            if combined:
                out_path = self.export_combined_path.get().strip()
                try:
                    write_dataframes_to_excel(combined, out_path)
                    self._log(f"[OK] Combined workbook written \u2192 {out_path}")
                except Exception as exc:
                    self._log(f"[ERROR] Could not write combined workbook: {exc}")
                    fail_count += 1

        def finish():
            self.progress.stop()
            self.progress.pack_forget()
            self.convert_btn.state(["!disabled"])
            messagebox.showinfo(
                "Conversion finished",
                f"Done!\n\n{ok_count} file(s) converted successfully."
                + (f"\n{fail_count} file(s) failed \u2014 see the log for details." if fail_count else "")
            )
        self.after(0, finish)


# ---------------------------------------------------------------------------
# TAB 2: IMPORT TO SQL
# ---------------------------------------------------------------------------


class BatteryImportTab(tk.Frame):
    """Import to SQL: connect to MySQL, organize Excel files into 'cells',
    and import them into database tables."""

    def __init__(self, parent):
        super().__init__(parent, bg=BG)

        self.engine_root = None
        self.engine_db = None
        self.host = None
        self.port = None
        self.user = None
        self.password = None
        self.selected_db = tk.StringVar(value="")
        self.selected_cell = tk.StringVar(value="")

        self._build_header()

        self.container = tk.Frame(self, bg=BG)
        self.container.pack(fill="both", expand=True, padx=24, pady=(0, 24))

        self.frames = {}
        for F in (ConnectionScreen, DatabaseScreen):
            frame = F(self.container, self)
            self.frames[F.__name__] = frame
            frame.place(x=0, y=0, relwidth=1, relheight=1)

        self.show_frame("ConnectionScreen")

    def _build_header(self):
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=24, pady=(20, 10))
        tk.Label(header, text="\U0001F50B Battery EIS Data Importer", bg=BG, fg=TEXT,
                 font=FONT_TITLE).pack(side="left")
        self.status_dot = tk.Label(header, text="\u25CF", bg=BG, fg=DANGER, font=("Segoe UI", 14))
        self.status_dot.pack(side="right")
        self.status_label = tk.Label(header, text="Not connected", bg=BG, fg=TEXT_MUTED, font=FONT_SMALL)
        self.status_label.pack(side="right", padx=(0, 8))

    def set_connected_status(self, connected: bool, detail: str = ""):
        if connected:
            self.status_dot.config(fg=ACCENT)
            self.status_label.config(text=detail or "Connected")
        else:
            self.status_dot.config(fg=DANGER)
            self.status_label.config(text="Not connected")

    def show_frame(self, name):
        self.frames[name].tkraise()
        if hasattr(self.frames[name], "on_show"):
            self.frames[name].on_show()


class ConnectionScreen(ttk.Frame):
    def __init__(self, parent, app: BatteryImportTab):
        super().__init__(parent, style="TFrame")
        self.app = app

        ttk.Label(self, text="Step 1 \u2014 Connect to a database server", style="Title.TLabel",
                  font=FONT_BOLD).pack(anchor="w", pady=(0, 4))
        ttk.Label(self, text="Choose a saved connection or enter a new one.",
                  style="TLabel", foreground=TEXT_MUTED).pack(anchor="w", pady=(0, 16))

        list_panel = tk.Frame(self, bg=PANEL, padx=16, pady=16)
        list_panel.pack(fill="x", pady=(0, 16))

        ttk.Label(list_panel, text="Available connections", style="Panel.TLabel",
                  font=FONT_BOLD).pack(anchor="w", pady=(0, 8))

        self.conn_list = tk.Listbox(list_panel, bg=PANEL_ALT, fg=TEXT, height=4,
                                     selectbackground=ACCENT_DARK, selectforeground="#06231c",
                                     borderwidth=0, highlightthickness=0, font=FONT,
                                     activestyle="none")
        self.conn_list.pack(fill="x")
        self.connections = [{"host": "localhost", "port": "3306", "user": "root"}]
        self._refresh_conn_list()

        ttk.Button(list_panel, text="Connect to selected", style="Accent.TButton",
                   command=self.connect_selected).pack(anchor="e", pady=(12, 0))

        add_panel = tk.Frame(self, bg=PANEL, padx=16, pady=16)
        add_panel.pack(fill="x")

        ttk.Label(add_panel, text="Add a new connection", style="Panel.TLabel",
                  font=FONT_BOLD).pack(anchor="w", pady=(0, 8))

        grid = tk.Frame(add_panel, bg=PANEL)
        grid.pack(fill="x")

        self.host_var = tk.StringVar(value="localhost")
        self.port_var = tk.StringVar(value="3306")
        self.user_var = tk.StringVar(value="root")

        self._labeled_entry(grid, "Host", self.host_var, 0)
        self._labeled_entry(grid, "Port", self.port_var, 1)
        self._labeled_entry(grid, "Username", self.user_var, 2)

        ttk.Button(add_panel, text="+ Add connection", style="Ghost.TButton",
                   command=self.add_connection).pack(anchor="e", pady=(12, 0))

        self.error_label = ttk.Label(self, text="", style="TLabel", foreground=DANGER)
        self.error_label.pack(anchor="w", pady=(12, 0))

    def _labeled_entry(self, parent, label, var, col):
        box = tk.Frame(parent, bg=PANEL)
        box.grid(row=0, column=col, padx=(0, 16), sticky="w")
        ttk.Label(box, text=label, style="Muted.TLabel").pack(anchor="w")
        ttk.Entry(box, textvariable=var, width=16).pack(anchor="w", pady=(2, 0))

    def _refresh_conn_list(self):
        self.conn_list.delete(0, tk.END)
        for c in self.connections:
            self.conn_list.insert(tk.END, f"  {c['user']}@{c['host']}:{c['port']}")
        if self.connections:
            self.conn_list.selection_set(0)

    def add_connection(self):
        host, port, user = self.host_var.get().strip(), self.port_var.get().strip(), self.user_var.get().strip()
        if not host or not port or not user:
            self.error_label.config(text="Host, port, and username are all required.")
            return
        self.connections.append({"host": host, "port": port, "user": user})
        self._refresh_conn_list()
        self.error_label.config(text="")

    def connect_selected(self):
        sel = self.conn_list.curselection()
        if not sel:
            self.error_label.config(text="Select a connection first.")
            return
        conn = self.connections[sel[0]]
        password = self._ask_password(conn["host"], conn["user"])
        if password is None:
            return

        self.error_label.config(text="Connecting\u2026")
        self.update_idletasks()

        try:
            url = f"mysql+pymysql://{conn['user']}:{password}@{conn['host']}:{conn['port']}/"
            engine = create_engine(url, connect_args={"connect_timeout": 6})
            with engine.connect() as c:
                c.execute(text("SELECT 1"))
        except Exception as exc:
            self.error_label.config(text=f"Connection failed: {exc}")
            return

        self.app.engine_root = engine
        self.app.host, self.app.port = conn["host"], conn["port"]
        self.app.user, self.app.password = conn["user"], password
        self.app.set_connected_status(True, f"{conn['user']}@{conn['host']}:{conn['port']}")
        self.error_label.config(text="")
        self.app.show_frame("DatabaseScreen")

    def _ask_password(self, host, user):
        dialog = tk.Toplevel(self)
        dialog.title("Enter password")
        dialog.configure(bg=PANEL)
        dialog.geometry("340x180")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        result = {"value": None}

        ttk.Label(dialog, text=f"Password for {user}@{host}", style="Panel.TLabel",
                  font=FONT_BOLD).pack(pady=(20, 4), padx=20, anchor="w")
        ttk.Label(dialog, text="Credentials are used only to connect and are not stored.",
                  style="Muted.TLabel", wraplength=300).pack(padx=20, anchor="w")

        pw_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=pw_var, show="\u2022", width=30)
        entry.pack(padx=20, pady=(12, 12))
        entry.focus_set()

        def submit(event=None):
            result["value"] = pw_var.get()
            dialog.destroy()

        def cancel():
            dialog.destroy()

        entry.bind("<Return>", submit)

        btns = tk.Frame(dialog, bg=PANEL)
        btns.pack(pady=(0, 10))
        ttk.Button(btns, text="Cancel", style="Ghost.TButton", command=cancel).pack(side="left", padx=6)
        ttk.Button(btns, text="Connect", style="Accent.TButton", command=submit).pack(side="left", padx=6)

        dialog.wait_window()
        return result["value"]


class DatabaseScreen(ttk.Frame):
    def __init__(self, parent, app: BatteryImportTab):
        super().__init__(parent, style="TFrame")
        self.app = app

        ttk.Label(self, text="Step 2 \u2014 Choose a database", style="Title.TLabel",
                  font=FONT_BOLD).pack(anchor="w", pady=(0, 4))
        self.subtitle = ttk.Label(self, text="", style="TLabel", foreground=TEXT_MUTED)
        self.subtitle.pack(anchor="w", pady=(0, 16))

        canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.body = tk.Frame(canvas, bg=BG)
        self.body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.body, anchor="nw", width=760)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        db_panel = tk.Frame(self.body, bg=PANEL, padx=16, pady=16)
        db_panel.pack(fill="x", pady=(0, 16))

        ttk.Label(db_panel, text="Available databases", style="Panel.TLabel",
                  font=FONT_BOLD).pack(anchor="w", pady=(0, 8))

        self.db_list = tk.Listbox(db_panel, bg=PANEL_ALT, fg=TEXT, height=5,
                                   selectbackground=ACCENT_DARK, selectforeground="#06231c",
                                   borderwidth=0, highlightthickness=0, font=FONT,
                                   activestyle="none")
        self.db_list.pack(fill="x")
        self.db_list.bind("<<ListboxSelect>>", self.on_db_selected)

        self.no_db_label = ttk.Label(db_panel, text="No databases found on this server.",
                                      style="Panel.TLabel", foreground=DANGER)

        row = tk.Frame(db_panel, bg=PANEL)
        row.pack(fill="x", pady=(12, 0))
        self.new_db_var = tk.StringVar(value="BATTERY_AES")
        ttk.Entry(row, textvariable=self.new_db_var, width=24).pack(side="left")
        ttk.Button(row, text="+ Create database", style="Ghost.TButton",
                   command=self.create_database).pack(side="left", padx=8)
        ttk.Button(row, text="\u21BB Refresh", style="Ghost.TButton",
                   command=self.refresh_databases).pack(side="left", padx=(0, 8))
        ttk.Button(row, text="\U0001F5D1 Delete selected", style="Danger.TButton",
                   command=self.delete_database).pack(side="left")

        self.cell_panel = tk.Frame(self.body, bg=PANEL, padx=16, pady=16)

        self.cell_target_label = ttk.Label(self.cell_panel, text="", style="Panel.TLabel",
                                            font=FONT_BOLD)
        self.cell_target_label.pack(anchor="w", pady=(0, 8))

        self.cell_list = tk.Listbox(self.cell_panel, bg=PANEL_ALT, fg=TEXT, height=5,
                                     selectbackground=ACCENT_DARK, selectforeground="#06231c",
                                     borderwidth=0, highlightthickness=0, font=FONT,
                                     activestyle="none")
        self.cell_list.pack(fill="x")
        self.cell_list.bind("<<ListboxSelect>>", self.on_cell_selected)

        self.no_cell_label = ttk.Label(self.cell_panel, text="No cells yet \u2014 create one below.",
                                        style="Panel.TLabel", foreground=TEXT_MUTED)

        cell_row = tk.Frame(self.cell_panel, bg=PANEL)
        cell_row.pack(fill="x", pady=(12, 0))
        self.new_cell_var = tk.StringVar(value="CELL_01")
        ttk.Entry(cell_row, textvariable=self.new_cell_var, width=24).pack(side="left")
        ttk.Button(cell_row, text="+ Create cell", style="Ghost.TButton",
                   command=self.create_cell).pack(side="left", padx=8)
        ttk.Button(cell_row, text="\U0001F5D1 Delete selected cell", style="Danger.TButton",
                   command=self.delete_cell).pack(side="left")

        self.files_panel = tk.Frame(self.body, bg=PANEL, padx=16, pady=16)

        self.files_target_label = ttk.Label(self.files_panel, text="", style="Panel.TLabel",
                                             font=FONT_BOLD)
        self.files_target_label.pack(anchor="w", pady=(0, 8))

        self.files_list = tk.Listbox(self.files_panel, bg=PANEL_ALT, fg=TEXT, height=6,
                                      selectbackground=ACCENT_DARK, selectforeground="#06231c",
                                      borderwidth=0, highlightthickness=0, font=FONT_SMALL,
                                      activestyle="none")
        self.files_list.pack(fill="x", pady=(0, 8))

        files_btn_row = tk.Frame(self.files_panel, bg=PANEL)
        files_btn_row.pack(fill="x")
        ttk.Button(files_btn_row, text="Select Excel files\u2026", style="Ghost.TButton",
                   command=self.select_files).pack(side="left")
        ttk.Button(files_btn_row, text="Remove selected file", style="Danger.TButton",
                   command=self.remove_selected_file).pack(side="left", padx=8)

        self.import_panel = tk.Frame(self.body, bg=PANEL, padx=16, pady=16)

        ttk.Label(self.import_panel, text="Step 4 \u2014 Import everything to the database",
                  style="Panel.TLabel", font=FONT_BOLD).pack(anchor="w", pady=(0, 8))
        self.import_target_label = ttk.Label(self.import_panel, text="", style="Muted.TLabel")
        self.import_target_label.pack(anchor="w", pady=(0, 8))

        self.import_btn = ttk.Button(self.import_panel, text="Import all cells to database",
                                      style="Accent.TButton", command=self.start_import)
        self.import_btn.pack(anchor="e")

        self.progress = ttk.Progressbar(self.import_panel, mode="indeterminate")

        self.log_box = tk.Text(self.import_panel, height=8, bg="#0c141c", fg=TEXT_MUTED,
                                font=("Consolas", 9), borderwidth=0, highlightthickness=0)
        self.log_box.pack(fill="both", expand=True, pady=(10, 0))
        self.log_box.configure(state="disabled")

    def on_show(self):
        self.subtitle.config(text=f"Connected as {self.app.user}@{self.app.host}:{self.app.port}")
        self.cell_panel.pack_forget()
        self.files_panel.pack_forget()
        self.import_panel.pack_forget()
        self.app.selected_db.set("")
        self.app.selected_cell.set("")
        self.refresh_databases()

    def refresh_databases(self):
        self.db_list.delete(0, tk.END)
        self.no_db_label.pack_forget()
        try:
            with self.app.engine_root.connect() as c:
                result = c.execute(text("SHOW DATABASES"))
                dbs = [row[0] for row in result if row[0] not in SYSTEM_DBS]
        except Exception as exc:
            messagebox.showerror("Error", f"Could not list databases:\n{exc}")
            return

        if not dbs:
            self.no_db_label.pack(anchor="w", pady=(8, 0))
        else:
            for db in dbs:
                self.db_list.insert(tk.END, f"  {db}")

    def create_database(self):
        name = self.new_db_var.get().strip()
        if not name:
            return
        try:
            with self.app.engine_root.connect() as c:
                c.execute(text(f"CREATE DATABASE IF NOT EXISTS `{name}`"))
                c.commit()
        except Exception as exc:
            messagebox.showerror("Error", f"Could not create database:\n{exc}")
            return
        self.refresh_databases()

    def delete_database(self):
        sel = self.db_list.curselection()
        if not sel:
            messagebox.showwarning("No database selected", "Select a database to delete first.")
            return
        db_name = self.db_list.get(sel[0]).strip()

        confirmed = messagebox.askyesno(
            "Delete database",
            f"Permanently delete database '{db_name}'?\n\n"
            "This drops it on the MySQL server and also removes its local "
            "cell folders and any Excel files stored in them. This cannot be undone."
        )
        if not confirmed:
            return

        try:
            with self.app.engine_root.connect() as c:
                c.execute(text(f"DROP DATABASE `{db_name}`"))
                c.commit()
        except Exception as exc:
            messagebox.showerror("Error", f"Could not delete database:\n{exc}")
            return

        delete_db_folder(db_name)

        if self.app.selected_db.get() == db_name:
            self.app.selected_db.set("")
            self.app.selected_cell.set("")
            self.app.engine_db = None
            self.cell_panel.pack_forget()
            self.files_panel.pack_forget()
            self.import_panel.pack_forget()

        self.refresh_databases()

    def on_db_selected(self, event=None):
        sel = self.db_list.curselection()
        if not sel:
            return
        db_name = self.db_list.get(sel[0]).strip()
        self.app.selected_db.set(db_name)
        self.app.selected_cell.set("")

        url = (f"mysql+pymysql://{self.app.user}:{self.app.password}@"
               f"{self.app.host}:{self.app.port}/{db_name}")
        self.app.engine_db = create_engine(url)

        self.cell_target_label.config(text=f"Cells in database: {db_name}")
        self.cell_panel.pack(fill="x", pady=(0, 16))
        self.files_panel.pack_forget()

        self.import_target_label.config(text=f"Importing into database: {db_name}")
        self.import_panel.pack(fill="both", expand=True)

        self.refresh_cells()

    def refresh_cells(self):
        db_name = self.app.selected_db.get()
        self.cell_list.delete(0, tk.END)
        self.no_cell_label.pack_forget()
        cells = list_cells(db_name)
        if not cells:
            self.no_cell_label.pack(anchor="w", pady=(8, 0))
        else:
            for c in cells:
                self.cell_list.insert(tk.END, f"  {c}")

    def create_cell(self):
        db_name = self.app.selected_db.get()
        if not db_name:
            messagebox.showwarning(
                "No database selected",
                "Select a database above first \u2014 every cell is created inside "
                "whichever database is currently selected."
            )
            return
        name = self.new_cell_var.get().strip()
        if not name:
            messagebox.showwarning("Cell name required", "Type a name for the cell first.")
            return
        create_cell(db_name, name)
        self.refresh_cells()

        for i in range(self.cell_list.size()):
            if self.cell_list.get(i).strip() == name:
                self.cell_list.selection_clear(0, tk.END)
                self.cell_list.selection_set(i)
                self.cell_list.see(i)
                self.on_cell_selected()
                break

    def delete_cell(self):
        db_name = self.app.selected_db.get()
        sel = self.cell_list.curselection()
        if not sel:
            messagebox.showwarning("No cell selected", "Select a cell to delete first.")
            return
        cell_name = self.cell_list.get(sel[0]).strip()

        confirmed = messagebox.askyesno(
            "Delete cell",
            f"Delete cell '{cell_name}'?\n\n"
            "This removes its local folder (and every Excel file inside it) "
            f"AND drops every table already imported from it in database "
            f"'{db_name}'. This cannot be undone."
        )
        if not confirmed:
            return

        delete_cell(db_name, cell_name)

        dropped = drop_tables_with_prefix(self.app.engine_db, f"{sanitize_name(cell_name)}__")

        if self.app.selected_cell.get() == cell_name:
            self.app.selected_cell.set("")
            self.files_panel.pack_forget()
        self.refresh_cells()

        if dropped:
            messagebox.showinfo(
                "Cell deleted",
                f"Deleted cell '{cell_name}' and dropped {dropped} table(s) from '{db_name}'."
            )

    def on_cell_selected(self, event=None):
        sel = self.cell_list.curselection()
        if not sel:
            return
        cell_name = self.cell_list.get(sel[0]).strip()
        self.app.selected_cell.set(cell_name)

        self.files_target_label.config(
            text=f"Excel files in '{self.app.selected_db.get()}' / '{cell_name}'")
        self.files_panel.pack(fill="x", pady=(0, 16), before=self.import_panel)
        self.refresh_files()

    def refresh_files(self):
        db_name = self.app.selected_db.get()
        cell_name = self.app.selected_cell.get()
        self.files_list.delete(0, tk.END)
        for f in list_cell_files(db_name, cell_name):
            self.files_list.insert(tk.END, f)

    def select_files(self):
        db_name = self.app.selected_db.get()
        cell_name = self.app.selected_cell.get()
        if not db_name or not cell_name:
            messagebox.showwarning("No cell selected", "Choose (or create) a cell first.")
            return

        files = filedialog.askopenfilenames(
            title="Select Excel files to add to this cell",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if files:
            copy_files_into_cell(db_name, cell_name, files)
            self.refresh_files()

    def remove_selected_file(self):
        db_name = self.app.selected_db.get()
        cell_name = self.app.selected_cell.get()
        sel = self.files_list.curselection()
        if not sel:
            messagebox.showwarning("No file selected", "Select a file to remove first.")
            return
        filename = self.files_list.get(sel[0]).strip()

        confirmed = messagebox.askyesno(
            "Remove file",
            f"Remove '{filename}' from cell '{cell_name}'?\n\n"
            "This deletes the local copy AND drops any tables already "
            f"imported from it in database '{db_name}'."
        )
        if not confirmed:
            return

        delete_cell_file(db_name, cell_name, filename)

        file_stub = sanitize_name(os.path.splitext(filename)[0])
        dropped = drop_tables_with_prefix(
            self.app.engine_db, f"{sanitize_name(cell_name)}__{file_stub}"
        )

        self.refresh_files()
        if dropped:
            messagebox.showinfo("File removed", f"Removed '{filename}' and dropped {dropped} table(s).")

    def _log(self, msg):
        def append():
            self.log_box.configure(state="normal")
            self.log_box.insert(tk.END, msg + "\n")
            self.log_box.see(tk.END)
            self.log_box.configure(state="disabled")
        self.after(0, append)

    def start_import(self):
        db_name = self.app.selected_db.get()
        if not db_name or self.app.engine_db is None:
            messagebox.showwarning("No database", "Select a database first.")
            return

        cells = list_cells(db_name)
        total_files = sum(len(list_cell_files(db_name, c)) for c in cells)
        if total_files == 0:
            messagebox.showwarning(
                "Nothing to import",
                "No Excel files found in any cell of this database yet."
            )
            return

        self.import_btn.state(["disabled"])
        self.progress.pack(fill="x", pady=(4, 8))
        self.progress.start(10)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", tk.END)
        self.log_box.configure(state="disabled")

        thread = threading.Thread(target=self._run_import, args=(db_name, cells), daemon=True)
        thread.start()

    def _run_import(self, db_name, cells):
        total_tables = 0
        for cell_name in cells:
            files = list_cell_files(db_name, cell_name)
            if not files:
                continue
            self._log(f"--- Cell '{cell_name}' ---")
            for fname in files:
                filepath = os.path.join(cell_dir(db_name, cell_name), fname)
                self._log(f"Importing {fname} \u2026")
                total_tables += import_excel_file(self.app.engine_db, filepath, self._log,
                                                   table_prefix=cell_name)

        def finish():
            self.progress.stop()
            self.progress.pack_forget()
            self.import_btn.state(["!disabled"])
            messagebox.showinfo("Upload successful",
                                 f"Upload successful!\n\n{total_tables} table(s) imported into "
                                 f"'{db_name}' from {len(cells)} cell(s).")
        self.after(0, finish)


# ---------------------------------------------------------------------------
# MAIN WINDOW: Chrome-style tabs merging both tools
# ---------------------------------------------------------------------------


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Tools Suite")
        self.geometry("900x800")
        self.configure(bg=BG)
        self.minsize(800, 700)

        build_shared_style(self)

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        dta_tab = DtaToExcelTab(notebook)
        sql_tab = BatteryImportTab(notebook)

        notebook.add(dta_tab, text="  DTA \u2192 Excel  ")
        notebook.add(sql_tab, text="  Import to SQL  ")


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()