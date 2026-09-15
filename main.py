from pathlib import Path
import pandas as pd


INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")

REQUIRED_COLUMNS = {
    "SKU",
    "Description",
    "Quantity",
    "Location",
}


def get_excel_files():
    """
    Returns all Excel files found in the input directory.
    """
    return sorted(INPUT_DIR.glob("*.xlsx"))


def clean_dataframe(df):
    """
    Cleans the dataframe and returns:
    - cleaned dataframe
    - number of duplicates removed
    """

    original_rows = len(df)

    # Clean column names
    df.columns = df.columns.str.strip()

    # Remove duplicate rows
    df = df.drop_duplicates()

    duplicates_removed = original_rows - len(df)

    # Clean text fields
    for column in df.select_dtypes(include="str").columns:
        df[column] = df[column].str.strip()

    return df, duplicates_removed


def validate_columns(df):
    """
    Checks if all required columns exist.
    """

    missing_columns = REQUIRED_COLUMNS - set(df.columns)

    return missing_columns


def analyze_dataframe(df):
    """
    Generates data quality metrics.
    """

    numeric_quantity = pd.to_numeric(
        df["Quantity"],
        errors="coerce",
    )

    return {
        "missing_sku": int(df["SKU"].isna().sum()),
        "missing_location": int(df["Location"].isna().sum()),
        "invalid_quantity": int(numeric_quantity.isna().sum()),
        "negative_quantity": int((numeric_quantity < 0).sum()),
    }

def process_file(input_file, save_individual=True):
    """
    Processes one Excel file.

    Returns a dictionary containing the cleaned dataframe
    and its data-quality metrics.
    """

    print()
    print(f"Processing: {input_file.name}")
    print("-" * 40)

    try:
        df = pd.read_excel(input_file)

    except Exception as exc:
        print(f"ERROR: Could not read file: {exc}")
        return None

    original_rows = len(df)

    # Clean column names before validation
    df.columns = df.columns.str.strip()

    # Validate required columns
    missing_columns = validate_columns(df)

    if missing_columns:
        print("ERROR: Missing required columns:")

        for column in sorted(missing_columns):
            print(f"- {column}")

        return None

    # Clean dataframe
    df, duplicates_removed = clean_dataframe(df)

    # Analyze data quality
    metrics = analyze_dataframe(df)

    # Output filenames
    stem = input_file.stem

    cleaned_file = OUTPUT_DIR / f"{stem}_cleaned.xlsx"
    report_file = OUTPUT_DIR / f"{stem}_report.xlsx"

    # Build quality report
    report = pd.DataFrame(
        {
            "Metric": [
                "Original Rows",
                "Cleaned Rows",
                "Duplicates Removed",
                "Missing SKU",
                "Missing Location",
                "Invalid Quantity",
                "Negative Quantity",
            ],
            "Value": [
                original_rows,
                len(df),
                duplicates_removed,
                metrics["missing_sku"],
                metrics["missing_location"],
                metrics["invalid_quantity"],
                metrics["negative_quantity"],
            ],
        }
    )

    # Save individual files only when requested
    if save_individual:
        df.to_excel(cleaned_file, index=False)
        report.to_excel(report_file, index=False)

    # Console report
    print(f"Original rows: {original_rows}")
    print(f"Cleaned rows: {len(df)}")
    print(f"Duplicates removed: {duplicates_removed}")
    print(f"Missing SKU: {metrics['missing_sku']}")
    print(f"Missing Location: {metrics['missing_location']}")
    print(f"Invalid Quantity: {metrics['invalid_quantity']}")
    print(f"Negative Quantity: {metrics['negative_quantity']}")

    if save_individual:
        print()
        print(f"Created: {cleaned_file}")
        print(f"Created: {report_file}")

    return {
        "filename": input_file.name,
        "dataframe": df,
        "original_rows": original_rows,
        "cleaned_rows": len(df),
        "duplicates_removed": duplicates_removed,
        **metrics,
    }

def process_all_files():
    """
    Processes every Excel file individually.
    """

    excel_files = get_excel_files()

    if not excel_files:
        print()
        print("No Excel files found in data/input.")
        return

    print()
    print(f"Files found: {len(excel_files)}")

    for input_file in excel_files:
        process_file(input_file)

    print()
    print("Processing completed.")


def consolidate_workbook():
    """
    Creates one Excel workbook:
    - one sheet per source file
    - one Summary sheet
    """

    excel_files = get_excel_files()

    if not excel_files:
        print()
        print("No Excel files found in data/input.")
        return

    results = []

    for input_file in excel_files:
        result = process_file(input_file,save_individual=False,)

        if result:
            results.append(result)

    if not results:
        print()
        print("No valid files were processed.")
        return

    consolidated_file = OUTPUT_DIR / "consolidated_report.xlsx"

    summary_rows = []

    with pd.ExcelWriter(
        consolidated_file,
        engine="openpyxl",
    ) as writer:

        for result in results:

            sheet_name = Path(
                result["filename"]
            ).stem[:31]

            result["dataframe"].to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
            )

            summary_rows.append(
                {
                    "File": result["filename"],
                    "Original Rows": result["original_rows"],
                    "Cleaned Rows": result["cleaned_rows"],
                    "Duplicates Removed": result[
                        "duplicates_removed"
                    ],
                    "Missing SKU": result["missing_sku"],
                    "Missing Location": result[
                        "missing_location"
                    ],
                    "Invalid Quantity": result[
                        "invalid_quantity"
                    ],
                    "Negative Quantity": result[
                        "negative_quantity"
                    ],
                }
            )

        summary_df = pd.DataFrame(summary_rows)

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )

    print()
    print(f"Consolidated workbook created: {consolidated_file}")

def combine_all_files():
    """
    Combines all valid cleaned Excel files
    into one worksheet.
    """

    excel_files = get_excel_files()

    if not excel_files:
        print()
        print("No Excel files found in data/input.")
        return

    combined_frames = []
    summary_rows = []

    expected_columns = None

    for input_file in excel_files:

        result = process_file(
            input_file,
            save_individual=False,
        )

        if not result:
            continue

        df = result["dataframe"].copy()

        current_columns = set(df.columns)

        if expected_columns is None:
            expected_columns = current_columns

        elif current_columns != expected_columns:
            print()
            print(
                f"WARNING: {input_file.name} has a "
                "different column structure."
            )
            print("File skipped.")
            continue

        # Keep track of source file
        df.insert(
            0,
            "Source File",
            input_file.name,
        )

        combined_frames.append(df)

        summary_rows.append(
            {
                "File": result["filename"],
                "Original Rows": result["original_rows"],
                "Cleaned Rows": result["cleaned_rows"],
                "Duplicates Removed": result[
                    "duplicates_removed"
                ],
                "Missing SKU": result["missing_sku"],
                "Missing Location": result[
                    "missing_location"
                ],
                "Invalid Quantity": result[
                    "invalid_quantity"
                ],
                "Negative Quantity": result[
                    "negative_quantity"
                ],
            }
        )

    if not combined_frames:
        print()
        print("No compatible files were available to combine.")
        return

    combined_df = pd.concat(
        combined_frames,
        ignore_index=True,
    )

    summary_df = pd.DataFrame(summary_rows)

    output_file = OUTPUT_DIR / "combined_data.xlsx"

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl",
    ) as writer:

        combined_df.to_excel(
            writer,
            sheet_name="Combined Data",
            index=False,
        )

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )

    print()
    print(f"Combined workbook created: {output_file}")
    print(f"Total combined rows: {len(combined_df)}")

def show_menu():
    """
    Displays the main application menu.
    """

    while True:

        print()
        print("DataClean Pro")
        print("=============")
        print("1. Process files separately")
        print("2. Consolidate files into separate sheets")
        print("3. Combine all files into one sheet")
        print("4. Exit")

        option = input("\nSelect an option: ").strip()

        if option == "1":
           process_all_files()

        elif option == "2":
           consolidate_workbook()

        elif option == "3":
           combine_all_files()

        elif option == "4":
           print()
           print("DataClean Pro closed.")
           break

        else:
           print()
           print("Invalid option. Please select 1, 2, 3 or 4.")


def main():
    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    show_menu()


if __name__ == "__main__":
    main()