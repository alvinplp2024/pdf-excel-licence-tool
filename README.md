# PDF → Excel Licence Tool

A desktop GUI tool that matches PDF filenames to rows in an Excel workbook and fills in a licence number column — one PDF at a time, with full user control over every change.

---

## What It Does

PDF files are named in the format `YREF-LICENCENO.pdf` (e.g. `Y24242025120097-2519946.pdf`).

The tool:
1. Reads your Excel workbook and lets you pick the worksheet, Y Ref column, and Licence No. column
2. Scans a folder of PDFs and parses each filename
3. Looks up the Y Ref in Excel and — only if the licence cell is **empty** — prompts you to fill it
4. Saves a new updated workbook and a matching report (the original is never overwritten)

---

## Requirements

- Python 3.8+
- `openpyxl` library

Install the dependency:

```bash
pip install openpyxl
```

---

## How to Run

### macOS / Linux

```bash
cd SCRIPT
python licence_tool.py
```

### Windows

Double-click `SCRIPT/run_windows.bat`

Or from a terminal:

```cmd
cd SCRIPT
python licence_tool.py
```

---

## Step-by-Step Usage

| Step | What you do |
|------|-------------|
| 1 | Browse and select your Excel workbook (`.xlsx` or `.xlsm`) |
| 2 | Choose the worksheet (auto-selects "Customer List" if found) |
| 3 | Map the Y Ref column and Licence No. column (auto-detected where possible) |
| 4 | Browse and select the folder containing your PDF files |
| 5 | Review the summary and click **START PROCESSING** |

For each PDF you will see one of these outcomes:

| Result | Meaning |
|--------|---------|
| `MATCH FOUND` | Y Ref found, licence cell is empty — you can fill it |
| `ALREADY FILLED` | Licence cell already has a value — skipped automatically |
| `NOT FOUND` | Y Ref does not exist in the Excel column |
| `DUPLICATE Y REF` | Y Ref appears on multiple rows — tool will not guess |
| `INVALID FILENAME` | Filename does not match the `YREF-LICENCE.pdf` pattern |

You can **CONTINUE**, **SKIP**, or **EXIT** at any point.

---

## PDF Filename Format

Files must follow this exact pattern:

```
<YREF>-<LICENCENO>.pdf
```

Examples:
```
Y24242025120097-2519946.pdf
Y24242026020015-2520047.pdf
```

---

## Output Files

After processing, clicking **SAVE RESULTS** produces two files saved next to your original workbook:

- `<workbook>_UPDATED_<timestamp>.xlsx` — updated workbook with licence numbers filled in
- `<workbook>_MATCHING_REPORT_<timestamp>.xlsx` — full audit report with a Results sheet and Summary sheet

The original workbook is **never modified**.

---

## Project Structure

```
AUT/
├── SCRIPT/
│   ├── licence_tool.py     # Main application
│   └── run_windows.bat     # Windows launcher
├── PDFs/                   # Sample/working PDF files
├── PDFs -folder/           # Additional PDF files
└── README.md
```

---

## Notes

- Existing licence values are **never overwritten** — the tool only fills empty cells
- Duplicate Y Refs in Excel are flagged and skipped
- The tool processes PDFs one at a time so you stay in control of every change
- No internet connection required — fully offline

---

## License

This project is not currently under an open-source license. All rights reserved.
