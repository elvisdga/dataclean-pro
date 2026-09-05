import pandas as pd


input_file = "data/input/sample.xlsx"
output_file = "data/output/cleaned.xlsx"
report_file = "data/output/data_quality_report.xlsx"


df = pd.read_excel(input_file)

print("Data loaded successfully.")

original_rows = len(df)
print(f"Original rows: {original_rows}")

# Remove duplicate rows
df = df.drop_duplicates()

cleaned_rows = len(df)
duplicates_removed = original_rows - cleaned_rows

print(f"Rows after removing duplicates: {cleaned_rows}")

# Clean column names
df.columns = df.columns.str.strip()

# Clean text fields
for column in df.select_dtypes(include="str").columns:
    df[column] = df[column].str.strip()

# Data quality checks
missing_sku = df["SKU"].isna().sum()
missing_location = df["Location"].isna().sum()
negative_quantity = (df["Quantity"] < 0).sum()

print()
print("Data Quality Report")
print("-------------------")
print(f"Duplicates removed: {duplicates_removed}")
print(f"Missing SKU: {missing_sku}")
print(f"Missing Location: {missing_location}")
print(f"Negative Quantity: {negative_quantity}")

# Save cleaned data
df.to_excel(output_file, index=False)

# Build quality report
report = pd.DataFrame(
    {
        "Metric": [
            "Original Rows",
            "Cleaned Rows",
            "Duplicates Removed",
            "Missing SKU",
            "Missing Location",
            "Negative Quantity",
        ],
        "Value": [
            original_rows,
            cleaned_rows,
            duplicates_removed,
            missing_sku,
            missing_location,
            negative_quantity,
        ],
    }
)

report.to_excel(report_file, index=False)

print()
print(f"Cleaned file created: {output_file}")
print(f"Quality report created: {report_file}")