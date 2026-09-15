# DataClean Pro

DataClean Pro is a Python automation tool for cleaning, validating, and consolidating Excel files.

## Current Features

- Process multiple Excel files automatically
- Remove duplicate rows
- Clean text fields
- Detect missing SKU values
- Detect missing locations
- Detect invalid quantities
- Detect negative quantities
- Generate cleaned Excel files
- Generate data quality reports
- Consolidate multiple Excel files into one workbook
- Combine multiple cleaned files into one sheet
- Generate summary reports

## Menu

```text
1. Process files separately
2. Consolidate files into separate sheets
3. Combine all files into one sheet
4. Exit


Project Structure
dataclean-pro/
├── data/
│   ├── input/
│   └── output/
├── main.py
├── generate_test_files.py
├── requirements.txt
├── .gitignore
└── README.md
Installation
Create and activate a virtual environment:

python -m venv .venv
source .venv/bin/activate
Install dependencies:

pip install -r requirements.txt
Usage
Place Excel files inside:

data/input/
Run:

python main.py
Then select one of the available processing options.

Output
Generated files are saved in:

data/output/
Technologies
Python 3.12

pandas

openpyxl

Project Status
Current version: v0.3

Planned improvements:

Flexible column mapping

Better validation

Configuration file

Improved error handling

More reusable project structure

Possible graphical or web interface
