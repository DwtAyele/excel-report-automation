"""Create a deliberately messy sample sales file for the demo."""
from pathlib import Path

import pandas as pd

rows = [
    ["2026-01-05", " North ", "Laptop", "2", "$850.00", "A. Bekele"],
    ["2026-01-05", "north", "Laptop", "2", "$850.00", "A. Bekele"],        # duplicate
    ["01/07/2026", "South", "monitor", "5", "199.99", "S. Tadesse"],
    ["2026-01-09", "EAST", "Keyboard", "10", "$45", "M. Haile"],
    ["2026-01-12", "West", "Mouse", "", "$25.50", "L. Girma"],              # missing qty
    ["2026-01-15", "South", "Laptop", "1", "$900.00", "S. Tadesse"],
    ["Jan 18 2026", "North", "Monitor", "3", "$210.00", "A. Bekele"],
    ["2026-01-20", "East", "laptop ", "2", "$870.00", "M. Haile"],
    ["2026-01-22", "West", "Keyboard", "4", "$47.00", "L. Girma"],
    ["2026-01-25", "", "Mouse", "8", "$24.00", "Unknown"],                  # missing region
    ["2026-02-02", "North", "Mouse", "6", "$25.00", "A. Bekele"],
    ["2026-02-03", "South", "Keyboard", "7", "$44.00", "S. Tadesse"],
    ["2026-02-08", "East", "Monitor", "2", "$205.00", "M. Haile"],
    ["2026-02-10", "West", "Laptop", "1", "$880.00", "L. Girma"],
    ["2026-02-14", "north", "Laptop", "3", "$860.00", "A. Bekele"],
]
df = pd.DataFrame(rows, columns=["Order Date", "Region", "Product", "Qty", "Unit Price", "Sales Rep"])
out = Path(__file__).parent / "sample_data" / "messy_sales.xlsx"
df.to_excel(out, index=False)
print(f"Wrote {out}")
