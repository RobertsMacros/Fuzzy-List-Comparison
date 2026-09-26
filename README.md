# Fuzzy List Compare

**Roberts Macros: no macro too micro.**

An Excel VBA macro that compares two lists and finds the items that are probably the same, even when they are not spelled identically. Leading numbering, punctuation, spacing and case are ignored. Each item in the first list is paired with its best match in the second list and given a confidence score.

It is useful for checking a list of documents against a folder of files, one index against another, or any two lists that should line up but were typed by different people.

## What you get

The macro adds a new sheet named `<first list> vs <second list>`. Your original sheet is never changed. The new sheet shows:

- every item from the first list, its best match from the second list and a confidence %, sorted by confidence
- the best-match column coloured by strength: dark green (95%+), green (80%+), yellow-green (60%+), yellow (40%+), orange (20%+), red (no match)
- items in the first list with no match above your threshold
- items in the second list that nothing matched

## Install

You need desktop Excel for Windows with macros enabled. Excel for Mac cannot build the window (UserForms cannot be edited on Mac), so set it up on Windows first; the finished workbook then runs on either.

1. Download `FuzzyListCompare.bas` and `frmFuzzyCompare.frm` from this repository.
2. Open the workbook you want to use it in, or your Personal Macro Workbook (`PERSONAL.XLSB`) to have it in every workbook.
3. Press **Alt+F11** to open the Visual Basic Editor.
4. Choose **File → Import File…** and import `FuzzyListCompare.bas`. Do the same for `frmFuzzyCompare.frm`.
5. The form file holds the code but not the on-screen controls, so add them once. Double-click `frmFuzzyCompare`, open the **Toolbox** and add these, setting each control's **(Name)** in the Properties window:

   | Control | (Name) | Notes |
   |---|---|---|
   | TextBox | `txtInstructions` | Set `MultiLine` to True and `ScrollBars` to Vertical |
   | ComboBox | `cmbPrimary` | The list to check |
   | ComboBox | `cmbSecondary` | The list to compare against |
   | ComboBox | `cmbThreshold` | The confidence threshold |
   | CommandButton | `btnImport` | Caption: Import Filenames |
   | CommandButton | `btnRun` | Caption: Run Comparison |
   | CommandButton | `btnClose` | Caption: Close |

6. Save the workbook as `.xlsm` (or save `PERSONAL.XLSB`).

To add a button, go to **File → Options → Customize Ribbon**, add a new group and add the `FuzzyListCompare` macro to it.

## Use

1. Put each list in its own column with a header in row 1 and the items from row 2 down. Lists can be in any columns on the same sheet.
2. Run the `FuzzyListCompare` macro (**Alt+F8**, choose it, **Run**).
3. In the window:
   - **Primary list**: the list you want to check.
   - **Secondary list**: the list to check it against.
   - **Threshold**: how close a match must be to count (default 70%). Lower it to catch looser matches; raise it to cut false matches.
4. Click **Run Comparison**. The results sheet opens.

### Comparing against a folder of files

Click **Import Filenames**, type the column letter to fill (for example `C`), then pick a folder. The file names are written into that column with the folder name as its header. You can import several folders into different columns before running a comparison. Windows only.

## Source files

| File | What it does |
|---|---|
| `FuzzyListCompare.bas` | Cleaning, scoring, colouring and the results sheet. `FuzzyListCompare` is the macro to run |
| `frmFuzzyCompare.frm` | The window: list pickers, threshold, filename import |
