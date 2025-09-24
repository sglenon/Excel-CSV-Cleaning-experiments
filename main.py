import os
import sys
import argparse
import glob
import traceback

# Import functions from src scripts
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))
from sheets_to_excel import separate_sheets_with_openpyxl
from preprocessing_excel_sheets import recalculate_and_refresh_sheets
from find_table_boundaries import find_table_boundaries
from process_with_pandas import process_table_with_pandas

def main():
    parser = argparse.ArgumentParser(
        description="Orchestrate Excel cleaning pipeline: split sheets, refresh, find table boundaries, and clean data."
    )
    parser.add_argument("input_excel_file", help="Path to the source Excel file (.xlsx).")
    parser.add_argument("output_directory", help="Directory where all outputs will be saved.")
    args = parser.parse_args()

    input_excel_file = args.input_excel_file
    output_dir = args.output_directory

    if not os.path.exists(input_excel_file):
        print(f"❌ Input file '{input_excel_file}' does not exist.")
        sys.exit(1)
    os.makedirs(output_dir, exist_ok=True)

    # Define subfolders for each step
    split_dir = os.path.join(output_dir, "split")
    refreshed_dir = os.path.join(output_dir, "refreshed")
    boundaries_dir = os.path.join(output_dir, "boundaries")
    cleaned_dir = os.path.join(output_dir, "cleaned")
    for d in [split_dir, refreshed_dir, boundaries_dir, cleaned_dir]:
        os.makedirs(d, exist_ok=True)

    print(f"\n[1/4] Splitting sheets from '{input_excel_file}' into '{split_dir}' ...")
    try:
        separate_sheets_with_openpyxl(input_excel_file, split_dir)
    except Exception as e:
        print(f"❌ Failed to split sheets: {e}")
        traceback.print_exc()
        sys.exit(1)

    # Find all generated sheet files
    base_name = os.path.splitext(os.path.basename(input_excel_file))[0]
    sheet_files = sorted(glob.glob(os.path.join(split_dir, f"{base_name}_sheet*.xlsx")))

    if not sheet_files:
        print("❌ No sheet files were generated. Exiting.")
        sys.exit(1)

    print(f"\n[2/4] Processing each sheet file ...")
    summary = []
    for sheet_file in sheet_files:
        try:
            print(f"\n--- Processing sheet file: {os.path.basename(sheet_file)} ---")

            # Step 2: Refresh and recalculate
            print("  [2.1] Refreshing formulas and data ...")
            refreshed_file = os.path.join(
                refreshed_dir, os.path.basename(sheet_file).replace(".xlsx", "_refreshed.xlsx")
            )
            # Copy the split file to refreshed_dir first, then refresh in place
            import shutil
            shutil.copy2(sheet_file, refreshed_file)
            recalculate_and_refresh_sheets(refreshed_file)

            # Step 3: Find table boundaries
            print("  [2.2] Finding table boundaries ...")
            boundaries_json = os.path.join(
                boundaries_dir, os.path.basename(sheet_file).replace(".xlsx", "_boundaries.json")
            )
            find_table_boundaries(refreshed_file, boundaries_json)

            # Step 4: Process with pandas
            print("  [2.3] Cleaning and saving final outputs ...")
            cleaned_excel = os.path.join(
                cleaned_dir, os.path.basename(sheet_file).replace(".xlsx", "_cleaned.xlsx")
            )
            cleaned_csv = os.path.join(
                cleaned_dir, os.path.basename(sheet_file).replace(".xlsx", "_cleaned.csv")
            )
            process_table_with_pandas(refreshed_file, boundaries_json, cleaned_excel, cleaned_csv)

            print(f"✅ Finished processing '{os.path.basename(sheet_file)}'.")
            summary.append((sheet_file, "Success", cleaned_excel, cleaned_csv))
        except Exception as e:
            print(f"❌ Error processing '{os.path.basename(sheet_file)}': {e}")
            traceback.print_exc()
            summary.append((sheet_file, "Failed", None, None))

    print("\n[3/4] Processing complete. Summary:")
    for entry in summary:
        sheet, status, excel, csv = entry
        print(f"  - {os.path.basename(sheet)}: {status}")
        if status == "Success":
            print(f"      Cleaned Excel: {os.path.basename(excel)}")
            print(f"      Cleaned CSV:   {os.path.basename(csv)}")

    failed = [s for s in summary if s[1] != "Success"]
    if failed:
        print("\n[4/4] Some sheets failed to process. See errors above.")
        sys.exit(2)
    else:
        print("\n[4/4] All sheets processed successfully.")

if __name__ == "__main__":
    main()