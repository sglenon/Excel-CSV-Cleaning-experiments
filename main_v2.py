"""
Main pipeline script (v2) for Excel/CSV cleaning and metadata extraction.

- Uses the new _v2.py modules for improved sheet naming, header cleaning, and metadata extraction.
- Ensures all outputs are PostgreSQL-friendly and metadata is saved for SQL ingestion.

Usage:
    python main_v2.py <input_excel_file.xlsx> <output_folder>

Dependencies:
    - openpyxl
    - pandas
    - xlwings
    - openai
    - dotenv
"""

import sys
import os

from src.sheets_to_excel_v2 import separate_sheets_with_openpyxl
from src.find_table_boundaries_v2 import find_table_boundaries
from src.process_with_pandas_v2 import process_table_with_pandas
from src.preprocessing_excel_sheets_v2 import recalculate_and_refresh_sheets

def main(input_excel, output_folder):
    # Step 1: Separate sheets with improved naming
    sheets_output_folder = os.path.join(output_folder, "separated_sheets")
    separate_sheets_with_openpyxl(input_excel, sheets_output_folder)

    # Step 2: For each separated sheet, find table boundaries, refresh, clean, and extract metadata
    for fname in os.listdir(sheets_output_folder):
        if not fname.endswith(".xlsx"):
            continue
        sheet_path = os.path.join(sheets_output_folder, fname)
        base = os.path.splitext(fname)[0]
        boundaries_json = os.path.join(output_folder, "boundaries", f"{base}_boundaries.json")
        refreshed_path = os.path.join(output_folder, "refreshed", f"{base}_refreshed.xlsx")
        cleaned_excel = os.path.join(output_folder, "cleaned", f"{base}_cleaned.xlsx")
        cleaned_csv = os.path.join(output_folder, "cleaned", f"{base}_cleaned.csv")
        metadata_json = os.path.join(output_folder, "metadata", f"{base}_metadata.json")

        os.makedirs(os.path.dirname(boundaries_json), exist_ok=True)
        os.makedirs(os.path.dirname(refreshed_path), exist_ok=True)
        os.makedirs(os.path.dirname(cleaned_excel), exist_ok=True)
        os.makedirs(os.path.dirname(metadata_json), exist_ok=True)

        # Find table boundaries
        find_table_boundaries(sheet_path, boundaries_json)

        # Refresh sheet (optional, can be skipped if not needed)
        recalculate_and_refresh_sheets(sheet_path)

        # Clean and extract metadata
        process_table_with_pandas(
            input_file=sheet_path,
            boundaries_json_path=boundaries_json,
            final_excel_path=cleaned_excel,
            final_csv_path=cleaned_csv,
            metadata_json_path=metadata_json
        )

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python main_v2.py <input_excel_file.xlsx> <output_folder>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])