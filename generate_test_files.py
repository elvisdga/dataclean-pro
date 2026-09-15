import pandas as pd
from pathlib import Path


input_dir = Path("data/input")

files = {
    "inventory_january.xlsx": {
        "SKU": ["A100", "A101", "A102", "A102"],
        "Description": ["Keyboard", "Mouse", "Monitor", "Monitor"],
        "Quantity": [10, 25, 8, 8],
        "Location": ["A-01", "A-02", "B-01", "B-01"],
    },
    "inventory_february.xlsx": {
        "SKU": ["B100", None, "B102", "B103"],
        "Description": ["Cable", "Adapter", "Dock", "Headset"],
        "Quantity": [15, 12, -4, 20],
        "Location": ["C-01", "C-02", "C-03", "C-04"],
    },
    "inventory_march.xlsx": {
        "SKU": ["C100", "C101", "C102"],
        "Description": ["Laptop", "Tablet", "Phone"],
        "Quantity": [5, 7, 9],
        "Location": ["D-01", None, "D-03"],
    },
}


for filename, data in files.items():
    df = pd.DataFrame(data)
    df.to_excel(input_dir / filename, index=False)

print("Test files created successfully.")
