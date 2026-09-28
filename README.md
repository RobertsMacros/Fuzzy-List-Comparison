# Fuzzy List Compare

**Roberts Macros: no macro too micro.**

An Excel VBA macro that compares two lists and finds the items that are probably the same, even when they are not spelled identically. List numbering ("1.", "(3)", "01_"), file extensions (".pdf", ".docx"), punctuation, spacing and case are ignored. Years and other long numbers are kept, so "2020 Budget" and "2021 Budget" are not treated as the same. Each item in the first list is paired with its best match in the second list and given a confidence score.

It is useful for checking a list of documents against a folder of files, one index against another, or any two lists that should line up but were typed by different people.

## What you get

The macro adds a new sheet named `<first list> vs <second list>`. Your original sheet is never changed. The new sheet shows:

- every item from the first list, its best match from the second list and a confidence %, sorted by confidence
- the best-match column coloured by strength: dark green (95%+), green (80%+), yellow-green (60%+), yellow (40%+), orange (20%+), red (no match)
- items in the first list with no match above your threshold
- items in the second list that nothing matched

## Install

You need desktop Excel with macros enabled. Set it up in Excel for Windows; the finished workbook also runs in Excel for Mac (except **Import Filenames**, which is Windows only).

1. Download `FuzzyListCompare.bas` and `frmFuzzyCompare.frm` from this repository.
2. Open the workbook you want to use it in, or your Personal Macro Workbook (`PERSONAL.XLSB`) to have it in every workbook.
3. Press **Alt+F11** to open the Visual Basic Editor.
4. Choose **File → Import File…** and import `FuzzyListCompare.bas`. Do the same for `frmFuzzyCompare.frm`. The window builds its own buttons and boxes when it opens, so there is nothing to draw.
5. Save the workbook as `.xlsm` (or save `PERSONAL.XLSB`).

To add a button, go to **File → Options → Customize Ribbon**, add a new group and add the `FuzzyListCompare` macro to it.

## Use

1. Put each list in its own column with a header in row 1 and the items from row 2 down. Lists can be in any columns on the same sheet.
2. Run the `FuzzyListCompare` macro (**Alt+F8**, choose it, **Run**).
3. In the window:
   - **Primary list**: the list you want to check.
   - **Secondary list**: the list to check it against.
   - **Threshold**: how close a match must be to count (40–90%, default 70%). Lower it to catch looser matches; raise it to cut false matches.
4. Click **Run Comparison** (or press Enter; Esc closes the window). The results sheet opens in the same workbook, next to your lists. Blank cells are skipped.

### How the score works

Each item is split into letter pairs within each word ("smith" → sm, mi, it, th). The score is the share of letter pairs the two items have in common, so word order hardly matters but a jumble of the same letters does not count as a match ("listen" vs "silent" scores 20%).

### Updating from an older copy

In the Visual Basic Editor, right-click `FuzzyListCompare` and `frmFuzzyCompare` → **Remove** (choose **No** when asked to export), then import the new files as in Install step 4.

### Comparing against a folder of files

Click **Import Filenames**, type the column letter to fill (for example `C`), then pick a folder. The file names are written into that column with the folder name as its header. You can import several folders into different columns before running a comparison. Windows only.

## Source files

| File | What it does |
|---|---|
| `FuzzyListCompare.bas` | Cleaning, scoring, colouring and the results sheet. `FuzzyListCompare` is the macro to run |
| `frmFuzzyCompare.frm` | The window: list pickers, threshold, filename import. Its controls are created in code |
