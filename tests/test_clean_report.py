import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from clean_report import clean, run  # noqa: E402


def raw(rows):
    return pd.DataFrame(rows, columns=["Order Date", "Region", "Product", "Qty", "Unit Price", "Sales Rep"], dtype=str)


def test_normalizes_and_removes_duplicates():
    good, rejected = clean(raw([
        ["2026-01-05", " North ", "laptop ", "2", "$850.00", "a. bekele"],
        ["2026-01-05", "north", "Laptop", "2", "850", "A. Bekele"],
    ]))
    assert len(good) == 1 and rejected.empty
    row = good.iloc[0]
    assert (row.Region, row.Product, row.Total) == ("North", "Laptop", 1700.0)


def test_bad_rows_go_to_review_with_reason():
    good, rejected = clean(raw([
        ["2026-01-12", "West", "Mouse", "", "$25.50", "L. Girma"],
        ["2026-01-25", "", "Mouse", "8", "$24.00", "Unknown"],
    ]))
    assert good.empty
    assert list(rejected.Reason) == ["missing qty", "missing region"]


def test_mixed_date_formats():
    good, _ = clean(raw([
        ["01/07/2026", "South", "Monitor", "5", "199.99", "S. Tadesse"],
        ["Jan 18 2026", "North", "Monitor", "3", "$210.00", "A. Bekele"],
    ]))
    assert list(good["Order Date"].dt.strftime("%Y-%m-%d")) == ["2026-01-07", "2026-01-18"]


def test_end_to_end(tmp_path):
    src = tmp_path / "in.csv"
    raw([["2026-02-02", "North", "Mouse", "6", "$25.00", "A. Bekele"]]).to_csv(src, index=False)
    out = tmp_path / "report.xlsx"
    stats = run(src, out)
    assert stats["revenue"] == 150.0
    assert pd.ExcelFile(out).sheet_names == ["By Region", "By Product", "By Month", "Clean Data", "Needs Review"]
