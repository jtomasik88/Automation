#!/usr/bin/env python3
"""JobConditions: fill in the Job Conditions and Modifiers in a pop-up form, then save them
to a formatted Excel spreadsheet.

Double-click this file. A window opens with one row per item:
  - Type: what you type in (each box also has a drop-down of common choices)
Click "Create Spreadsheet", pick where to save, and the .xlsx is built for you.

Requires openpyxl:  py -m pip install openpyxl
"""

import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing package",
        "This script needs the openpyxl package.\n\n"
        "Open Command Prompt and run:\n\n    py -m pip install openpyxl",
    )
    sys.exit(1)

TITLE = "Job Conditions and Modifiers"
HEADERS = ("Item", "Type")

# (Item, drop-down suggestions for Type)
FIELDS = [
    ("Shape", ["Gable", "Single Slope"]),
    ("Width", []),
    ("Length", []),
    ("Eave Height", []),
    ("Eave", ["Gutter w/ trim", "Eave trim"]),
    ("Base", ["PR sq, cut w/ base trim (1/2\"), base girt"]),
    ("Roof Slope", ["1/2:12", "1:12", "2:12", "3:12", "4:12"]),
    ("Ridge Height", []),
    ("Purlins", ["8\"", "10\"", "12\""]),
    ("Girts", ["8 1/2\"", "10\"", "12\""]),
    ("Roof Panel", ["SSR", "PR"]),
    ("Roof Panel Gage", ["22", "24", "26"]),
    ("Roof Panel Color", []),
    ("Wall Panel Type", ["PR"]),
    ("Wall Panel Gage", ["24", "26"]),
    ("Wall Panel Color", []),
    ("Liner Panel", ["PR"]),
    ("Liner Panel Gage", ["24", "26"]),
    ("Liner Panel Color", []),
    ("Trim Color", []),
    ("Door Color", []),
    ("Downspout Color", []),
]

# Items whose Type is stored as a number (so Excel treats it like one).
NUMERIC_ITEMS = {"Roof Panel Gage", "Wall Panel Gage", "Liner Panel Gage"}


def as_cell_value(item, text):
    text = text.strip()
    if item in NUMERIC_ITEMS:
        try:
            return int(text)
        except ValueError:
            pass
    return text or None


def build_workbook(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Job Conditions"

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    title = ws.cell(row=1, column=1, value=TITLE)
    title.font = Font(size=14)
    title.alignment = Alignment(horizontal="center", vertical="center")
    for col in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = PatternFill("solid", fgColor="00B0F0")
        cell.border = border
    ws.row_dimensions[1].height = 22

    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font = Font(bold=True)
        cell.border = border

    for r, (item, type_) in enumerate(rows, start=3):
        for col, value in enumerate((item, as_cell_value(item, type_)), start=1):
            cell = ws.cell(row=r, column=col, value=value)
            cell.border = border

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = max(40, *(len(t) + 2 for _, t in rows))
    return wb


class JobConditionsForm:
    def __init__(self, root):
        self.root = root
        root.title(TITLE)
        root.resizable(False, False)

        frame = ttk.Frame(root, padding=12)
        frame.grid()

        ttk.Label(frame, text=TITLE, font=("Segoe UI", 13, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 8)
        )
        for col, header in enumerate(HEADERS):
            ttk.Label(frame, text=header, font=("Segoe UI", 10, "bold")).grid(
                row=1, column=col, sticky="w", padx=4
            )

        self.entries = []
        for r, (item, choices) in enumerate(FIELDS, start=2):
            ttk.Label(frame, text=item).grid(row=r, column=0, sticky="w", padx=4, pady=1)
            type_box = ttk.Combobox(frame, values=choices, width=42)
            type_box.grid(row=r, column=1, padx=4, pady=1)
            self.entries.append((item, type_box))

        buttons = ttk.Frame(frame)
        buttons.grid(row=len(FIELDS) + 2, column=0, columnspan=2, pady=(10, 0))
        ttk.Button(buttons, text="Create Spreadsheet", command=self.save).grid(row=0, column=0, padx=6)
        ttk.Button(buttons, text="Clear", command=self.clear).grid(row=0, column=1, padx=6)
        ttk.Button(buttons, text="Close", command=root.destroy).grid(row=0, column=2, padx=6)

        self.entries[0][1].focus_set()

    def clear(self):
        for _, type_box in self.entries:
            type_box.set("")
        self.entries[0][1].focus_set()

    def save(self):
        rows = [(item, box.get()) for item, box in self.entries]
        if not any(t.strip() for _, t in rows):
            messagebox.showwarning(TITLE, "Fill in at least one Type before saving.")
            return

        path = filedialog.asksaveasfilename(
            title="Save spreadsheet as",
            initialdir=Path(__file__).resolve().parent,
            initialfile=f"{TITLE}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return

        try:
            build_workbook(rows).save(path)
        except PermissionError:
            messagebox.showerror(
                TITLE, f"Couldn't save:\n{path}\n\nIf the file is open in Excel, close it and try again."
            )
            return

        if messagebox.askyesno(TITLE, f"Saved:\n{path}\n\nOpen it now?"):
            if hasattr(os, "startfile"):
                os.startfile(path)


def main():
    root = tk.Tk()
    JobConditionsForm(root)
    root.mainloop()


if __name__ == "__main__":
    main()
