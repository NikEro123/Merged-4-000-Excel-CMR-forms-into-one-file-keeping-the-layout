# CMR merge

Script I wrote at N.N.M. Group (logistics) to merge two old Excel workbooks full of CMR forms into one file.

## The problem

- Two .xls files with 4,000+ sheets, one CMR form per sheet
- Most forms were in both files
- Some sheets were empty
- Text was typed in AcadNusx, an old Georgian font. On computers without it, "თარიღი" shows up as "TariRi"

## What the script does

- Reads every sheet with its formatting (borders, colors, merged cells, column widths, row heights)
- Skips empty sheets and duplicates (a CMR is a duplicate if its key fields match one already copied)
- Converts AcadNusx text to real Georgian letters and uses Sylfaen, so it shows correctly everywhere
- Only converts the old-font parts of a cell, so mixed text like "MAXTHERM LTD" + a Georgian address stays correct
- Writes each unique CMR as its own sheet in one .xlsx, looking the same as the original form

## Result

2,197 unique CMRs in one file, 1,958 duplicates and 19 empty sheets skipped. Takes about 4 minutes.

## Run

    pip install xlrd openpyxl

Put `merge_cmr_forms.py` in the same folder as both .xls files and run it.
It finds the files by name (`11...xls` and `IEL...2024.xls`) and saves `CMR_merged.xlsx` next to them.

The data files aren't in this repo because they contain company and driver information.
