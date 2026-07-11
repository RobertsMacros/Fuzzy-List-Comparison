# Fuzzy List Comparison

Compare two lists in an Excel workbook for fuzzy duplicate matches. Results
are written to a new dedicated sheet, sorted by match confidence, with the
best-match column colour-coded. The original data is never modified.

This is a Python port of the original Excel VBA macro (`FuzzyListCompare.bas`
+ `frmFuzzyCompare.frm`).

## Requirements

- Python 3.10+
- [openpyxl](https://openpyxl.readthedocs.io/) (`pip install -r requirements.txt`)
- Tkinter (bundled with most Python installs) — only needed for the GUI

## Usage

### GUI

```
python fuzzy_compare_gui.py [workbook.xlsx]
```

1. Open a workbook (or pass its path on the command line) and pick a sheet.
2. Optionally click **Import Filenames** to load the file names from a folder
   into a column — the folder name becomes the column header. You can import
   multiple folders into different columns.
3. Select the **primary** list (the list you want to check) and the
   **secondary** list (the list to compare against).
4. Choose a confidence threshold (default 70%) and click **Run Comparison**.

### Command line

```
python fuzzy_list_compare.py workbook.xlsx --primary A --secondary B
python fuzzy_list_compare.py workbook.xlsx --primary "List A" --secondary "List B" --threshold 80
```

Lists can be referenced by column letter or by header name. Use `--sheet` to
pick a source sheet (defaults to the active sheet) and `--output` to write
results to a copy instead of saving in place.

## What it does

Each item is cleaned (leading numbering stripped, punctuation removed,
lowercased) and scored against every item in the secondary list using a
character-bag similarity measure. A new results sheet contains:

- The primary list sorted by best-match confidence (descending), with the
  best match and a colour-coded confidence score:
  - **Green** — strong match (80%+)
  - **Yellow** — moderate match (40–79%)
  - **Orange/Red** — weak or no match
- Items in the primary list with no match above the threshold
- Items in the secondary list that nothing matched

## Layout

- `fuzzy_list_compare.py` — core comparison logic and CLI (port of `FuzzyListCompare.bas`)
- `fuzzy_compare_gui.py` — Tkinter GUI (port of the `frmFuzzyCompare` UserForm)
