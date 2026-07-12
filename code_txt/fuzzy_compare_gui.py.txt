"""
frmFuzzyCompare - GUI front end for FuzzyListCompare
====================================================
Tkinter port of the original VBA UserForm (frmFuzzyCompare.frm).

Opens an .xlsx workbook, detects list columns (row-1 headers), lets the
user pick a primary and secondary list plus a confidence threshold, and
runs the fuzzy comparison. Also supports importing filenames from a
folder into a column, so folder contents can be compared as lists.

Usage:
    python fuzzy_compare_gui.py [workbook.xlsx]
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from fuzzy_list_compare import (
    cell_text,
    detect_list_columns,
    last_row_in_column,
    run_fuzzy_comparison,
)

INSTRUCTIONS = """\
FUZZY LIST COMPARE
-----------------------------------

Compare two lists for fuzzy duplicate matches.

IMPORT FILENAMES:
Click 'Import Filenames' to load file names from a folder into a \
column. You can import multiple folders into different columns before \
running the comparison.

HOW TO USE:
1. Select your PRIMARY list - the list you want to check.
2. Select the SECONDARY list - the list to compare against.
3. Choose a confidence threshold (default 70%).
4. Click 'Run Comparison'.

WHAT IT DOES:
Creates a new results sheet sorted by match confidence.
Each item shows its best match and a confidence %.

The best-match column is colour-coded:
  Green  = strong match (80%+)
  Yellow = moderate match (40-79%)
  Orange/Red = weak or no match

Below the main table:
- Primary items with no good match
- Secondary items not matched

Your original data is never modified.
"""

THRESHOLD_OPTIONS = ["50%", "60%", "70%", "80%", "90%"]
DEFAULT_THRESHOLD_INDEX = 2  # 70%


class FuzzyCompareApp:
    def __init__(self, root: tk.Tk, workbook_path: str | None = None):
        self.root = root
        self.root.title("Fuzzy List Compare")
        self.root.resizable(False, False)

        self.workbook = None
        self.workbook_path = None
        self.source_ws = None
        self.columns = []  # list of (column_index, header)

        self._build_ui()

        if workbook_path:
            self._load_workbook(workbook_path)

    # -------------------------------------------------------
    # UI construction
    # -------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}
        frame = ttk.Frame(self.root, padding=10)
        frame.grid(sticky="nsew")

        # Instructions
        txt = tk.Text(frame, width=64, height=20, wrap="word")
        txt.insert("1.0", INSTRUCTIONS)
        txt.configure(state="disabled")
        txt.grid(row=0, column=0, columnspan=2, **pad)

        # Workbook / sheet selection
        wb_frame = ttk.Frame(frame)
        wb_frame.grid(row=1, column=0, columnspan=2, sticky="ew", **pad)
        ttk.Button(wb_frame, text="Open Workbook...",
                   command=self.open_workbook).pack(side="left")
        self.lbl_workbook = ttk.Label(wb_frame, text="No workbook open")
        self.lbl_workbook.pack(side="left", padx=8)

        ttk.Label(frame, text="Sheet:").grid(row=2, column=0, sticky="e", **pad)
        self.cmb_sheet = ttk.Combobox(frame, state="disabled", width=40)
        self.cmb_sheet.grid(row=2, column=1, sticky="w", **pad)
        self.cmb_sheet.bind("<<ComboboxSelected>>", self._on_sheet_change)

        # List selection
        ttk.Label(frame, text="Primary list:").grid(row=3, column=0, sticky="e", **pad)
        self.cmb_primary = ttk.Combobox(frame, state="disabled", width=40)
        self.cmb_primary.grid(row=3, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Secondary list:").grid(row=4, column=0, sticky="e", **pad)
        self.cmb_secondary = ttk.Combobox(frame, state="disabled", width=40)
        self.cmb_secondary.grid(row=4, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Threshold:").grid(row=5, column=0, sticky="e", **pad)
        self.cmb_threshold = ttk.Combobox(frame, state="readonly", width=10,
                                          values=THRESHOLD_OPTIONS)
        self.cmb_threshold.current(DEFAULT_THRESHOLD_INDEX)
        self.cmb_threshold.grid(row=5, column=1, sticky="w", **pad)

        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=6, column=0, columnspan=2, **pad)
        self.btn_import = ttk.Button(btn_frame, text="Import Filenames",
                                     command=self.import_filenames,
                                     state="disabled")
        self.btn_import.pack(side="left", padx=4)
        self.btn_run = ttk.Button(btn_frame, text="Run Comparison",
                                  command=self.run_comparison, state="disabled")
        self.btn_run.pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Close",
                   command=self.root.destroy).pack(side="left", padx=4)

    # -------------------------------------------------------
    # Workbook handling
    # -------------------------------------------------------
    def open_workbook(self):
        path = filedialog.askopenfilename(
            title="Open workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if not path:
            return
        self._load_workbook(path)

    def _load_workbook(self, path: str):
        if not os.path.exists(path):
            if messagebox.askyesno(
                    "Create workbook",
                    f"{path} does not exist.\nCreate a new empty workbook?"):
                wb = Workbook()
                wb.save(path)
            else:
                return
        try:
            keep_vba = path.lower().endswith(".xlsm")
            self.workbook = load_workbook(path, keep_vba=keep_vba)
        except Exception as exc:
            messagebox.showerror("Fuzzy List Compare",
                                 f"Could not open workbook:\n{exc}")
            return

        self.workbook_path = path
        self.lbl_workbook.config(text=os.path.basename(path))

        self.cmb_sheet.config(state="readonly", values=self.workbook.sheetnames)
        self.cmb_sheet.current(self.workbook.index(self.workbook.active))
        self._on_sheet_change()
        self.btn_import.config(state="normal")

    def _on_sheet_change(self, _event=None):
        self.source_ws = self.workbook[self.cmb_sheet.get()]
        self.refresh_combos()

    # -------------------------------------------------------
    # refresh_combos - scan sheet headers, repopulate combos
    # -------------------------------------------------------
    def refresh_combos(self):
        pri_idx = self.cmb_primary.current()
        sec_idx = self.cmb_secondary.current()

        self.columns = detect_list_columns(self.source_ws)
        display = [f"{get_column_letter(col)} - {header}"
                   for col, header in self.columns]

        state = "readonly" if display else "disabled"
        self.cmb_primary.config(state=state, values=display)
        self.cmb_secondary.config(state=state, values=display)
        self.cmb_primary.set("")
        self.cmb_secondary.set("")

        # Restore selections if still valid
        if 0 <= pri_idx < len(display):
            self.cmb_primary.current(pri_idx)
        if 0 <= sec_idx < len(display):
            self.cmb_secondary.current(sec_idx)

        self.btn_run.config(state="normal" if len(self.columns) >= 2 else "disabled")

    # -------------------------------------------------------
    # import_filenames - import filenames from a folder
    # -------------------------------------------------------
    def import_filenames(self):
        # Step 1: Get target column letter
        col_letter = simpledialog.askstring(
            "Import Filenames",
            "Enter column letter to import into (e.g. A, B, C):",
            parent=self.root)
        if not col_letter:
            return
        col_letter = col_letter.strip().upper()

        try:
            target_col = column_index_from_string(col_letter)
        except ValueError:
            messagebox.showwarning("Import Filenames",
                                   f"Invalid column letter: {col_letter}")
            return

        # Step 2: Open folder picker
        folder_path = filedialog.askdirectory(
            title="Select folder containing files to import")
        if not folder_path:
            return

        # Step 3: Read filenames
        try:
            file_names = sorted(
                name for name in os.listdir(folder_path)
                if os.path.isfile(os.path.join(folder_path, name)))
        except OSError as exc:
            messagebox.showerror("Import Filenames",
                                 f"Could not read folder:\n{exc}")
            return

        if not file_names:
            messagebox.showwarning("Import Filenames",
                                   "No files found in the selected folder.")
            return

        # Step 4: Confirm overwrite if column has data
        ws = self.source_ws
        existing_hdr = cell_text(ws.cell(row=1, column=target_col).value).strip()
        if existing_hdr != "":
            if not messagebox.askyesno(
                    "Import Filenames",
                    f"Column {col_letter} already has data "
                    f"(header: {existing_hdr}).\n"
                    "Clear it and import filenames?"):
                return

        # Step 5: Clear column and write data
        last_row = max(last_row_in_column(ws, target_col), 1)
        for row in range(1, last_row + 1):
            ws.cell(row=row, column=target_col).value = None

        # Folder name as header, filenames in rows 2+
        folder_name = os.path.basename(os.path.normpath(folder_path))
        ws.cell(row=1, column=target_col, value=folder_name)
        for r, name in enumerate(file_names, start=2):
            ws.cell(row=r, column=target_col, value=name)

        if not self._save_workbook():
            return

        # Step 6: Refresh combos and notify
        self.refresh_combos()
        messagebox.showinfo(
            "Import Filenames",
            f"Imported {len(file_names)} filenames into column {col_letter}.\n"
            f"Header set to: {folder_name}")

    # -------------------------------------------------------
    # run_comparison - validate inputs and run comparison
    # -------------------------------------------------------
    def run_comparison(self):
        pri_idx = self.cmb_primary.current()
        sec_idx = self.cmb_secondary.current()

        if pri_idx == -1:
            messagebox.showwarning("Fuzzy List Compare",
                                   "Please select a primary list.")
            return
        if sec_idx == -1:
            messagebox.showwarning("Fuzzy List Compare",
                                   "Please select a secondary list.")
            return
        if pri_idx == sec_idx:
            messagebox.showwarning("Fuzzy List Compare",
                                   "Primary and secondary lists must be different.")
            return

        threshold = int(self.cmb_threshold.get().rstrip("%")) / 100

        primary_col, primary_header = self.columns[pri_idx]
        secondary_col, secondary_header = self.columns[sec_idx]

        try:
            result = run_fuzzy_comparison(
                self.workbook, self.source_ws, primary_col, secondary_col,
                primary_header, secondary_header, threshold)
        except ValueError as exc:
            messagebox.showwarning("Fuzzy List Compare", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("FuzzyListCompare Error",
                                 f"An error occurred:\n\n{exc}")
            return

        if not self._save_workbook():
            return

        messagebox.showinfo(
            "FuzzyListCompare",
            "Comparison complete!\n\n"
            f"Primary list: {primary_header} "
            f"({result['primary_count']} items)\n"
            f"Secondary list: {secondary_header} "
            f"({result['secondary_count']} items)\n\n"
            f"Results are on sheet: {result['sheet_name']}")
        self.root.destroy()

    def _save_workbook(self) -> bool:
        try:
            self.workbook.save(self.workbook_path)
            return True
        except Exception as exc:
            messagebox.showerror(
                "Fuzzy List Compare",
                f"Could not save workbook:\n{exc}\n\n"
                "If the file is open in Excel, close it and try again.")
            return False


def main():
    workbook_path = sys.argv[1] if len(sys.argv) > 1 else None
    root = tk.Tk()
    FuzzyCompareApp(root, workbook_path)
    root.mainloop()


if __name__ == "__main__":
    main()
