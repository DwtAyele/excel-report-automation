# Excel Report Automation

Turn a messy sales spreadsheet into a clean, formatted Excel report in one command.

This is the kind of task many teams still do by hand every week: fixing names, deleting duplicates, fixing dates and prices, then building summary tables. This script does it in seconds and flags the rows a person still needs to check.

## What it does

| Problem in the raw file | What the script does |
| --- | --- |
| `" North "`, `north`, `NORTH` | Trims and standardizes text |
| Same order entered twice | Removes exact duplicates |
| `2026-01-05`, `01/07/2026`, `Jan 18 2026` | Reads all date styles |
| `$850.00`, `850`, `$45` | Converts prices to numbers |
| Missing quantity or region | Moves the row to a **Needs Review** sheet with the reason |

The output workbook has five sheets: **By Region**, **By Product**, **By Month**, **Clean Data** and **Needs Review**. Each has styled headers, money formatting, frozen header rows and auto-sized columns.

## Example result (from the sample file)

```
Done: {'rows_in': 15, 'clean': 12, 'duplicates_removed': 1, 'needs_review': 2, 'revenue': 10935.95}
```

| Region | Orders | Units | Revenue |
| --- | --- | --- | --- |
| North | 4 | 14 | $5,060.00 |
| East | 3 | 14 | $2,600.00 |
| South | 3 | 13 | $2,207.95 |
| West | 2 | 5 | $1,068.00 |

## Run it

```bash
pip install -r requirements.txt
python make_sample_data.py                     # creates sample_data/messy_sales.xlsx
python clean_report.py sample_data/messy_sales.xlsx output/sales_report.xlsx
```

Works with `.xlsx` and `.csv` input on Windows, macOS and Linux.

## Tests

```bash
python -m pytest -q
```

Tests cover text cleanup, duplicate removal, mixed date formats, rejected rows and the full end-to-end report.

## Tech

Python, pandas, openpyxl, pytest.

---
Need something like this for your own files? I build custom Python automation: Excel and CSV reports, web scraping, file tasks and test automation.
