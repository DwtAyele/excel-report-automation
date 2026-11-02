"""Clean a messy sales spreadsheet and build a formatted Excel report.

Usage:
    python clean_report.py sample_data/messy_sales.xlsx output/sales_report.xlsx
"""
import argparse
import logging
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

log = logging.getLogger("clean_report")


def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (clean_rows, rejected_rows)."""
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # Normalize text columns
    for col in ["Region", "Product", "Sales Rep"]:
        df[col] = df[col].astype("string").str.strip().str.title()
    df["Region"] = df["Region"].replace("", pd.NA)

    # Parse dates written in different styles
    df["Order Date"] = pd.to_datetime(df["Order Date"].astype(str).str.strip(), errors="coerce", format="mixed")

    # Numbers: strip currency symbols and commas
    df["Unit Price"] = pd.to_numeric(
        df["Unit Price"].astype(str).str.replace(r"[$,\s]", "", regex=True), errors="coerce"
    )
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce")

    # Rows that cannot be used go to a separate sheet, with a reason
    reasons = pd.Series("", index=df.index, dtype="string")
    reasons[df["Order Date"].isna()] += "bad date; "
    reasons[df["Region"].isna()] += "missing region; "
    reasons[df["Qty"].isna()] += "missing qty; "
    reasons[df["Unit Price"].isna()] += "bad price; "
    bad = reasons != ""
    rejected = df[bad].assign(Reason=reasons[bad].str.rstrip("; "))
    good = df[~bad].drop_duplicates()

    good = good.assign(
        Qty=good["Qty"].astype(int),
        Total=(good["Qty"] * good["Unit Price"]).round(2),
        Month=good["Order Date"].dt.strftime("%Y-%m"),
    ).sort_values("Order Date")
    return good.reset_index(drop=True), rejected.reset_index(drop=True)


def summarize(good: pd.DataFrame) -> dict[str, pd.DataFrame]:
    by_region = good.groupby("Region", as_index=False).agg(Orders=("Total", "size"), Units=("Qty", "sum"), Revenue=("Total", "sum"))
    by_product = good.groupby("Product", as_index=False).agg(Units=("Qty", "sum"), Revenue=("Total", "sum"))
    by_month = good.pivot_table(index="Month", columns="Region", values="Total", aggfunc="sum", fill_value=0).reset_index()
    return {
        "By Region": by_region.sort_values("Revenue", ascending=False),
        "By Product": by_product.sort_values("Revenue", ascending=False),
        "By Month": by_month,
    }


def style_workbook(path: Path) -> None:
    wb = load_workbook(path)
    header_fill = PatternFill("solid", fgColor="1F3A5F")
    thin = Side(style="thin", color="CCCCCC")
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = Border(bottom=thin)
                header = ws.cell(row=1, column=cell.column).value
                if header in ("Unit Price", "Total", "Revenue") or (ws.title == "By Month" and cell.column > 1):
                    cell.number_format = '"$"#,##0.00'
                elif header == "Order Date":
                    cell.number_format = "yyyy-mm-dd"
        for col in range(1, ws.max_column + 1):
            letter = get_column_letter(col)
            width = max(len(str(c.value)) if c.value is not None else 0 for c in ws[letter])
            ws.column_dimensions[letter].width = min(max(width + 3, 10), 40)
        ws.freeze_panes = "A2"
    wb.save(path)


def run(src: Path, dest: Path) -> dict:
    raw = pd.read_excel(src, dtype=str) if src.suffix.lower() in (".xlsx", ".xls") else pd.read_csv(src, dtype=str)
    good, rejected = clean(raw)
    dest.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(dest, engine="openpyxl") as xw:
        for name, table in summarize(good).items():
            table.to_excel(xw, sheet_name=name, index=False)
        good.to_excel(xw, sheet_name="Clean Data", index=False)
        rejected.to_excel(xw, sheet_name="Needs Review", index=False)
    style_workbook(dest)
    stats = {"rows_in": len(raw), "clean": len(good), "duplicates_removed": len(raw) - len(rejected) - len(good),
             "needs_review": len(rejected), "revenue": float(good["Total"].sum())}
    log.info("Done: %s", stats)
    return stats


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input", type=Path, help="messy .xlsx or .csv file")
    p.add_argument("output", type=Path, help="report .xlsx to create")
    args = p.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    run(args.input, args.output)


if __name__ == "__main__":
    main()
