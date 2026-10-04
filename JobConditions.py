#!/usr/bin/env python3
"""JobConditions: fill in the Job Conditions and Modifiers in a pop-up form, then save them
to a formatted Excel spreadsheet.

Double-click this file. A window opens with one row per item:
  - Type: what you type in (each box also has a drop-down of common choices)
Click "Create Spreadsheet", pick where to save, and the .xlsx is built for you.

Uses only the Python standard library -- nothing extra to install.
"""

import os
import tkinter as tk
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape
from tkinter import filedialog, messagebox, ttk

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


def _cell(ref, value, style):
    if value is None:
        return f'<c r="{ref}" s="{style}"/>'
    if isinstance(value, int):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'


# Cell styles (index into cellXfs below): 0 default, 1 title, 2 header, 3 bordered body
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3"><font><sz val="11"/><name val="Aptos Narrow"/></font><font><sz val="14"/><name val="Aptos Narrow"/></font><font><b/><sz val="11"/><name val="Aptos Narrow"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF00B0F0"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="4"><xf/><xf fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf fontId="2" borderId="1" applyFont="1" applyBorder="1"/><xf borderId="1" applyBorder="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

STATIC_PARTS = {
    "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>""",
    "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
    "xl/workbook.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Job Conditions" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
    "xl/_rels/workbook.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""",
    "xl/styles.xml": STYLES_XML,
}


def write_xlsx(path, rows):
    """Write the Job Conditions sheet (title bar, headers, one row per item) to an .xlsx file."""
    sheet_rows = [
        f'<row r="1" ht="22" customHeight="1">{_cell("A1", TITLE, 1)}{_cell("B1", None, 1)}</row>',
        f'<row r="2">{_cell("A2", HEADERS[0], 2)}{_cell("B2", HEADERS[1], 2)}</row>',
    ]
    for r, (item, type_) in enumerate(rows, start=3):
        sheet_rows.append(f'<row r="{r}">{_cell(f"A{r}", item, 3)}{_cell(f"B{r}", as_cell_value(item, type_), 3)}</row>')

    type_width = max(40, *(len(t) + 2 for _, t in rows))
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="{type_width}" customWidth="1"/></cols>'
        f'<sheetData>{"".join(sheet_rows)}</sheetData>'
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
        "</worksheet>"
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in STATIC_PARTS.items():
            zf.writestr(name, data)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


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
            write_xlsx(path, rows)
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
