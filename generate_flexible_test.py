import pandas as pd
from pathlib import Path


input_dir = Path("data/input")

data = {
    "Item Number": ["X100", "X101", "X102"],
    "Product Name": ["Keyboard", "Mouse", "Monitor"],
    "Qty": [10, 20, 5],
    "Bin": ["A-01", "A-02", "B-01"],
}

df = pd.DataFrame(data)

output_file = input_dir / "flexible_columns_test.xlsx"

df.to_excel(output_file, index=False)

print(f"Created: {output_file}")
