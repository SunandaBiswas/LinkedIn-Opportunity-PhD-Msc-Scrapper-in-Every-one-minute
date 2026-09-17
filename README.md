# LinkedIn Funded PhD & Masters Postings — Master Workbook Builder

Consolidates per-country "funded PhD" and "funded Masters" LinkedIn job-posting
workbooks into a single master Excel file, and tracks how many entries are new
every time you rebuild it.

## Overview

Each country's postings live in their own small Excel workbook. This tool
scans a folder for every one of those workbooks, pulls their rows into one
combined file with a sheet per country/round, and builds a dashboard sheet on
top that totals everything up — including how many postings are new since the
last time you ran it. Any genuinely new postings are also saved as their own
small snapshot file per country/round, so you can check exactly what's new
without opening the full master workbook.

It's a pure local-file tool: it never logs into LinkedIn, never needs a
password, and only ever reads `.xlsx` files that already exist in the folder.

## Files in this repo

| File | Purpose |
|---|---|
| `Build_Master_Workbook.bat` | Double-click entry point (Windows). Checks for Python/`openpyxl`, then runs the builder. |
| `build_master_workbook.py` | The actual builder script. |
| `runner.txt` | Plain-language how-to-run guide and troubleshooting notes. |
| `<Country>_University_PhD_LinkedIn_Postings.xlsx` | One per country — funded PhD postings. |
| `<Country>_University_Masters_Funded_LinkedIn_Postings.xlsx` | One per country — funded Masters postings. |
| `All_Countries_Funded_PhD_Masters_Master.xlsx` | Generated output — the combined master workbook. Rebuilt on every run, don't hand-edit it. |
| `New_Entries/<Country>_<Round>_New_Entries_<date>.xlsx` | Generated output — one small snapshot file per country/round, created only when that run found genuinely new postings. |

Each per-country source workbook must contain a sheet named `All Postings`
with these 13 columns:

`No.` · `University / Institution` · `Location` · `Department / Lab` ·
`Category` · `Subfield` · `Position Summary` · `Start / Deadline` ·
`Contact` · `LinkedIn Post Link` · `Posted By` · `Source Type` · `Posted`

## Requirements

- Python 3
- [`openpyxl`](https://pypi.org/project/openpyxl/) — installed automatically
  by `Build_Master_Workbook.bat` if it isn't already present

## Usage

### Windows (recommended)

1. Make sure `All_Countries_Funded_PhD_Masters_Master.xlsx` is closed in
   Excel.
2. Double-click `Build_Master_Workbook.bat`.
3. Read the summary printed in the console window (see below), then open
   `All_Countries_Funded_PhD_Masters_Master.xlsx`.

It's always safe to re-run — it just re-reads whatever `.xlsx` files are
currently in the folder.

### Command line (any OS)

```bash
# Build from the folder the script lives in
python build_master_workbook.py

# Or point it at a specific folder
python build_master_workbook.py "/path/to/folder"
```

## Sample output

```
Scanning: D:\LinkedInSunanda
  +  India PhD                    12 rows   (was  10,  +2)   <- India_University_PhD_LinkedIn_Postings.xlsx
  +  Germany Masters                5 rows   (was   5,  +0)   <- Germany_University_Masters_Funded_LinkedIn_Postings.xlsx

Previous total entries : 55
New entries this run    : +3
Total entries now       : 58

Saved: D:\LinkedInSunanda\All_Countries_Funded_PhD_Masters_Master.xlsx
```

The same previous/new/total figures are also written into the "Master
Summary" sheet inside the generated workbook, so they're on record even
after the console window closes.

## Notes

- The Master Summary sheet totals (per-country and grand total) are live
  Excel formulas (`COUNTA`/`COUNTIF`), so they stay accurate even if you
  filter or sort a country's sheet afterward.
- The previous/new comparison tracks row *counts* per country/round, not
  individual postings — a removed posting offset by one added posting in
  the same country will show as no change.
- If this repo is public, consider whether the `.xlsx` data files should be
  committed as-is or excluded via `.gitignore`, since they contain scraped
  posting details (contacts, links) collected from LinkedIn.
