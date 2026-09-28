#!/usr/bin/env python3
"""
CA Report - Reverse TEXTJOIN Tool
=================================

Takes a "TextJoin" style CA Audit Report (.xlsx) - where each compliance
finding appears as ONE row with all affected hosts joined into a single
comma-separated "Host" cell - and expands it back into the normal
per-host format (one row per finding + host pair).

ANY report workbook can be uploaded: the report sheet, the header row and
the column positions are auto-detected from the header labels, so the
sheet does not have to be called "CA_Report" and the columns do not have
to be in the exact template order.

HOW TO RUN
----------
1. Make sure Python 3 is installed (python.org) - version 3.8 or newer.
2. Install the one required library (only needed once):
        pip install openpyxl
3. Double-click this file, or run from a terminal:
        python CA_TextJoin_Reverser.py
4. In the window that opens, click "Browse..." to choose your TextJoin
   .xlsx file, confirm/adjust the Save-As location, then click "Convert".

Or run from CLI:
        python CA_TextJoin_Reverser.py input.xlsx -o output.xlsx
"""

import copy
import os
import re
import sys
import threading
import traceback
from collections import Counter

try:
    from openpyxl import load_workbook
except ImportError:
    print("ERROR: The 'openpyxl' package is required.\n"
          "Install it with:  pip install openpyxl")
    sys.exit(1)

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ----------------------------------------------------------------------
# Conversion logic
# ----------------------------------------------------------------------

HEADER_ROW = 13
FIRST_DATA_ROW = 14
SHEET_NAME = "CA_Report"
SUMMARY_SHEET = "Summary"

# The report table is auto-detected, so these are only fallback positions.
COL_SR = 1
COL_TITLE = 2
COL_HOST = 3
COL_DESC = 4
COL_SOLUTION = 5
COL_STATUS = 6
LAST_COL = 6

STATUS_ORDER = {"FAILED": 0, "WARNING": 1, "PASSED": 2}

MAX_HEADER_SCAN = 40

NO_TABLE_MESSAGE = (
    "No CA report table found in the uploaded file. Expected a row of headers "
    "containing a title column, a host column and a status column."
)

# Header labels that identify each field of a CA report. A report only has
# to use one of these labels for a column to be recognised.
ROLE_KEYS = (
    ("sr", ("sr", "sr no", "s no", "s.no", "sr.no", "serial", "serial no", "#")),
    ("title", ("title", "control", "check", "finding", "policy", "control title",
               "control name", "issue", "plugin", "plugin name")),
    ("host", ("host", "host ip", "host name", "hostname", "affected host", "affected hosts",
              "ip address", "ip", "asset", "target")),
    ("desc", ("description", "details", "observation")),
    ("solution", ("solution", "remediation", "resolution", "recommendation")),
    ("status", ("status", "result", "compliance status", "state")),
)

# At least one of these must be present, otherwise the table is assumed to be
# a VA report and this tool leaves it alone.
CA_SPECIFIC = ("solution", "status")


def _norm(value):
    if value is None:
        return ""
    return " ".join(str(value).split()).lower()


def _label_matches(value, keys):
    text = _norm(value)
    if not text:
        return False
    return any(text == k or text.startswith(k) for k in keys)


def _read_roles(ws, row, max_col):
    """Map field name -> column index for the header labels on the given row."""
    roles = {}
    for c in range(1, max_col + 1):
        value = ws.cell(row, c).value
        for role, keys in ROLE_KEYS:
            if role not in roles and _label_matches(value, keys):
                roles[role] = c
                break
    return roles


def _find_report_table(wb):
    """Locate the CA report table in any workbook.

    Returns (score, worksheet, header_row, roles, last_col). Raises ValueError
    when no sheet looks like a CA report.
    """
    best = None
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        max_row = min(ws.max_row or 1, MAX_HEADER_SCAN)
        max_col = ws.max_column or 1
        for r in range(1, max_row + 1):
            roles = _read_roles(ws, r, max_col)
            if "title" not in roles or "host" not in roles:
                continue
            if not any(role in roles for role in CA_SPECIFIC):
                continue
            score = len(roles)
            if ws.title.strip().lower() == SHEET_NAME.lower():
                score += 2
            if r == HEADER_ROW:
                score += 1
            if best is None or score > best[0]:
                best = (score, ws, r, roles, max_col)
    if best is None:
        raise ValueError(NO_TABLE_MESSAGE)
    return best


def detect_score(input_path):
    """Return how strongly this workbook looks like a CA report (0 = not at all)."""
    try:
        wb = load_workbook(input_path, data_only=True)
    except Exception:
        return 0
    try:
        return _find_report_table(wb)[0]
    except ValueError:
        return 0
    finally:
        wb.close()


def _find_summary_sheet(wb):
    for name in wb.sheetnames:
        if name.strip().lower() == SUMMARY_SHEET.lower():
            return wb[name]
    for ws in wb.worksheets:
        if "summary" in ws.title.lower():
            return ws
    return None


def _find_last_data_row(ws, title_col, first_data_row):
    last = first_data_row - 1
    r = first_data_row
    empty_streak = 0
    max_row = ws.max_row or first_data_row
    while r <= max_row:
        title = ws.cell(r, title_col).value
        if title not in (None, ""):
            last = r
            empty_streak = 0
        else:
            empty_streak += 1
            if empty_streak > 5:
                break
        r += 1
    return last


def _copy_cell_style(src_cell, dst_cell):
    dst_cell.font = copy.copy(src_cell.font)
    dst_cell.border = copy.copy(src_cell.border)
    dst_cell.fill = copy.copy(src_cell.fill)
    dst_cell.number_format = copy.copy(src_cell.number_format)
    dst_cell.protection = copy.copy(src_cell.protection)
    dst_cell.alignment = copy.copy(src_cell.alignment)


def convert(input_path, output_path, progress_cb=None):
    def report(msg):
        if progress_cb:
            progress_cb(msg)

    report("Opening workbook...")
    wb = load_workbook(input_path, data_only=False)

    ws, header_row, roles, max_col = _find_report_table(wb)[1:]
    first_data_row = header_row + 1
    report(f"Found report table on sheet '{ws.title}' (headers on row {header_row}).")

    # Columns are addressed by the header labels that were actually found. A
    # column that was not recognised is left untouched rather than guessed at,
    # so an unfamiliar layout can never be overwritten.
    col_sr = roles.get("sr")
    col_title = roles["title"]
    col_host = roles["host"]
    col_desc = roles.get("desc")
    col_solution = roles.get("solution")
    col_status = roles.get("status")
    last_col = max(max_col, max(roles.values()))

    for key, col in (("Sr. no", col_sr), ("Description", col_desc),
                     ("Solution", col_solution), ("Status", col_status)):
        if col is None:
            report(f"Note: no '{key}' column found - leaving that column as it is.")

    last_row = _find_last_data_row(ws, col_title, first_data_row)
    if last_row < first_data_row:
        raise ValueError("No finding rows found to expand.")

    report(f"Reading {last_row - first_data_row + 1} joined rows...")

    known_cols = {c for c in (col_sr, col_title, col_host, col_desc, col_solution,
                              col_status) if c is not None}
    other_cols = [c for c in range(1, last_col + 1) if c not in known_cols]

    records = []
    for r in range(first_data_row, last_row + 1):
        title = ws.cell(r, col_title).value
        if title in (None, ""):
            continue
        host_raw = ws.cell(r, col_host).value
        desc = ws.cell(r, col_desc).value if col_desc else None
        solution = ws.cell(r, col_solution).value if col_solution else None
        status = ws.cell(r, col_status).value if col_status else None
        extras = {c: ws.cell(r, c).value for c in other_cols}

        hosts = []
        if host_raw is not None:
            for h in re.split(r"[,;\n\r]+", str(host_raw)):
                h = h.strip()
                if h:
                    hosts.append(h)
        if not hosts:
            hosts = [host_raw]

        for h in hosts:
            records.append({
                "title": title, "host": h, "desc": desc,
                "solution": solution, "status": status, "extras": extras,
            })

    report(f"Expanded to {len(records)} rows. Sorting...")

    def sort_key(rec):
        host = rec["host"] if rec["host"] is not None else ""
        status_rank = STATUS_ORDER.get(rec["status"], 99)
        title = rec["title"] if rec["title"] is not None else ""
        return (str(host), status_rank, str(title).lower())

    records.sort(key=sort_key)

    for i, rec in enumerate(records, start=1):
        rec["sr"] = i

    old_count = last_row - first_data_row + 1
    new_count = len(records)
    template_row = first_data_row

    if new_count > old_count:
        n_extra = new_count - old_count
        report(f"Inserting {n_extra} extra rows...")
        insert_at = last_row + 1
        ws.insert_rows(insert_at, amount=n_extra)
        template_height = ws.row_dimensions[template_row].height
        for offset in range(n_extra):
            new_r = insert_at + offset
            ws.row_dimensions[new_r].height = template_height
            for c in range(1, last_col + 1):
                _copy_cell_style(ws.cell(template_row, c), ws.cell(new_r, c))
    elif new_count < old_count:
        n_remove = old_count - new_count
        report(f"Removing {n_remove} unused rows...")
        remove_at = first_data_row + new_count
        ws.delete_rows(remove_at, amount=n_remove)

    report("Writing expanded rows...")
    for i, rec in enumerate(records):
        r = first_data_row + i
        for col, key in ((col_sr, "sr"), (col_title, "title"), (col_host, "host"),
                         (col_desc, "desc"), (col_solution, "solution"),
                         (col_status, "status")):
            if col is not None:
                ws.cell(r, col).value = rec[key]
        for col, value in rec["extras"].items():
            ws.cell(r, col).value = value

    stats = {"old_count": old_count, "new_count": new_count, "status_counts": {}}

    status_counts = Counter(rec["status"] for rec in records if rec["status"])
    stats["status_counts"] = dict(status_counts)

    sws = _find_summary_sheet(wb)
    if sws is not None:
        report("Updating Summary sheet counts...")
        for r in range(1, sws.max_row + 1):
            label = sws.cell(r, 5).value
            if isinstance(label, str) and label.strip() in ("FAILED", "WARNING", "PASSED"):
                sws.cell(r, 6).value = status_counts.get(label.strip(), 0)
            elif isinstance(label, str) and label.strip() == "Grand Total":
                sws.cell(r, 6).value = sum(status_counts.values())

    report("Saving output file...")
    wb.save(output_path)
    report("Done.")
    return stats


# ----------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CA Report - Reverse TEXTJOIN Tool")
        self.geometry("640x480")
        self.resizable(False, False)

        pad = {"padx": 12, "pady": 6}

        tk.Label(
            self, text="CA Report — Reverse TEXTJOIN Tool",
            font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", **pad)

        tk.Label(
            self,
            text="Upload any CA report and get back the non-TextJoin,\n"
                 "one-row-per-host format, ready to download and use directly.",
            justify="left", fg="#444"
        ).pack(anchor="w", padx=12)

        # --- Input file row ---
        frm_in = tk.Frame(self)
        frm_in.pack(fill="x", **pad)
        tk.Label(frm_in, text="Report file:", width=14, anchor="w").pack(side="left")
        self.input_var = tk.StringVar()
        tk.Entry(frm_in, textvariable=self.input_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(frm_in, text="Browse...", command=self.browse_input).pack(side="left")

        # --- Output file row ---
        frm_out = tk.Frame(self)
        frm_out.pack(fill="x", **pad)
        tk.Label(frm_out, text="Save as:", width=14, anchor="w").pack(side="left")
        self.output_var = tk.StringVar()
        tk.Entry(frm_out, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(frm_out, text="Browse...", command=self.browse_output).pack(side="left")

        # --- Convert button ---
        self.convert_btn = tk.Button(
            self, text="Convert", font=("Segoe UI", 11, "bold"),
            bg="#2563eb", fg="white", activebackground="#1d4ed8",
            command=self.start_convert
        )
        self.convert_btn.pack(pady=10)

        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=12)

        # --- Log box ---
        tk.Label(self, text="Log:").pack(anchor="w", padx=12, pady=(10, 0))
        self.log_box = scrolledtext.ScrolledText(self, height=14, state="disabled", font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    def log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def browse_input(self):
        path = filedialog.askopenfilename(
            title="Select the report .xlsx file",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.input_var.set(path)
            base, ext = os.path.splitext(path)
            suggested = base
            for suffix in ("_TextJoin", "TextJoin"):
                if suggested.endswith(suffix):
                    suggested = suggested[: -len(suffix)]
                    break
            if not suggested.endswith("_Normal") and suggested == base:
                suggested = base + "_Normal"
            self.output_var.set(suggested + ext)

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save expanded report as...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.output_var.set(path)

    def start_convert(self):
        in_path = self.input_var.get().strip()
        out_path = self.output_var.get().strip()

        if not in_path or not os.path.isfile(in_path):
            messagebox.showerror("Missing file", "Please select a valid report .xlsx file.")
            return
        if not out_path:
            messagebox.showerror("Missing output", "Please choose where to save the result.")
            return
        if os.path.abspath(in_path) == os.path.abspath(out_path):
            messagebox.showerror(
                "Same file",
                "Please choose a different Save-As location so the original file isn't overwritten."
            )
            return

        self.convert_btn.configure(state="disabled")
        self.progress.start(12)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        thread = threading.Thread(target=self._run_convert, args=(in_path, out_path), daemon=True)
        thread.start()

    def _run_convert(self, in_path, out_path):
        try:
            stats = convert(in_path, out_path, progress_cb=lambda m: self.after(0, self.log, m))
            self.after(0, self._on_success, stats, out_path)
        except Exception as e:
            tb = traceback.format_exc()
            self.after(0, self._on_error, str(e), tb)

    def _on_success(self, stats, out_path):
        self.progress.stop()
        self.convert_btn.configure(state="normal")
        sc = stats.get("status_counts", {})
        summary = (
            f"\nExpanded {stats['old_count']} joined rows -> {stats['new_count']} per-host rows.\n"
            f"Status breakdown: FAILED={sc.get('FAILED',0)}  WARNING={sc.get('WARNING',0)}\n"
            f"Saved to:\n{out_path}"
        )
        self.log(summary)
        messagebox.showinfo("Done", f"Conversion complete!\n\nSaved to:\n{out_path}")

    def _on_error(self, msg, tb):
        self.progress.stop()
        self.convert_btn.configure(state="normal")
        self.log(f"\nERROR: {msg}\n{tb}")
        messagebox.showerror("Conversion failed", msg)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="CA Report - Reverse TEXTJOIN Tool")
    parser.add_argument("input", nargs="?", help="Input TextJoin .xlsx file path")
    parser.add_argument("-o", "--output", help="Output file path (default: input_Normal.xlsx)")
    args = parser.parse_args()

    if args.input:
        in_path = args.input
        if args.output:
            out_path = args.output
        else:
            base, ext = os.path.splitext(in_path)
            out_path = base + "_Normal" + ext
        stats = convert(in_path, out_path, progress_cb=print)
        sc = stats.get("status_counts", {})
        print(f"\nExpanded {stats['old_count']} joined rows -> {stats['new_count']} per-host rows.")
        print(f"Status: FAILED={sc.get('FAILED',0)} WARNING={sc.get('WARNING',0)}")
        print(f"Saved to: {out_path}")
    else:
        app = App()
        app.mainloop()
