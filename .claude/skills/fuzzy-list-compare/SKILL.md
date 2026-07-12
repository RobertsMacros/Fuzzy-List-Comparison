---
name: fuzzy-list-compare
description: Compare two lists in an Excel workbook for fuzzy duplicate matches and produce a colour-coded results sheet. Use when the user wants to find duplicates, near-duplicates, or matching items between two columns/lists (e.g. filenames vs a catalogue, two inventories, two name lists), or asks to run a fuzzy comparison on a spreadsheet.
---

# Fuzzy List Compare

Compare a PRIMARY list against a SECONDARY list in an .xlsx workbook. Every
primary item is scored against every secondary item with a character-bag
similarity measure (0–100%), and a new results sheet is written with each
item's best match, colour-coded by confidence. The source data is never
modified.

## Requirements

- Python 3.10+ with `openpyxl` (`pip install -r requirements.txt` from the
  repository root, or `pip install openpyxl`).

## How to run

Preferred: the command-line interface in `fuzzy_list_compare.py` at the
repository root.

```bash
python fuzzy_list_compare.py <workbook.xlsx> --primary <col> --secondary <col> [options]
```

- `--primary` / `--secondary` accept a column letter (`A`) or a row-1 header
  name (`"List A"`). Lists live in columns: header in row 1, items in rows 2+.
- `--threshold N` — match threshold percentage, default 70.
- `--sheet NAME` — source sheet, default is the active sheet.
- `--output PATH` — write to a copy instead of saving in place.

Example:

```bash
python fuzzy_list_compare.py inventory.xlsx --primary "Warehouse" --secondary "Catalogue" --threshold 80
```

For programmatic use, import `run_fuzzy_comparison` from
`fuzzy_list_compare` (see the module docstring). A Tkinter GUI is also
available for interactive use on the user's machine: `python fuzzy_compare_gui.py`.

## Workflow

1. If the user's lists are not already in a workbook (e.g. plain text lists,
   CSV, or folder contents), first write them into an .xlsx: one column per
   list, header in row 1, items in rows 2+. To compare folder contents, write
   the filenames into a column with the folder name as the header.
2. Run the CLI. It exits 0 on success and prints the name of the results
   sheet; validation problems (empty list, bad column, same column twice)
   exit 1 with a message on stderr.
3. Report the summary to the user: item counts, the results sheet name, and
   how many items had no match above the threshold in each direction.

## Reading the results sheet

- Row 1: title with the lists compared and the threshold used.
- Row 3: headers; rows 4+: primary items sorted by confidence (descending)
  with their best secondary match. The best-match cell fill encodes
  confidence: dark green ≥95%, green ≥80%, yellow-green ≥60%, yellow ≥40%,
  orange ≥20%, red below.
- Two sections after the table list primary items with no match above the
  threshold and secondary items nothing matched (both red-filled).

## Notes

- Strings are cleaned before scoring: leading numbering stripped (`1.`,
  `(2)`, `3-`), punctuation removed, spaces collapsed, lowercased — so
  `apple_pie.txt` matches `1. Apple Pie` strongly.
- Similarity is order-insensitive (a character bag), so anagram-like strings
  score high; treat scores as candidates for human review, not proof.
- Re-running the same comparison adds a new numbered sheet (`... 2`, `... 3`)
  rather than overwriting.
