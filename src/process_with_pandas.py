# script_b_process_with_pandas.py

import pandas as pd
import json
from collections import defaultdict

HEADER_ROW_COUNT = 2

def process_table_with_pandas(input_file: str, boundaries_json_path: str, final_excel_path: str, final_csv_path: str):
    """
    Reads the original Excel file and uses the AI-found boundaries to perform
    a definitive, in-memory cleaning and structuring process with pandas.
    """
    print("\n --- Processing Table with Pandas-First Approach ---")

    # --- Step 1: Load Boundaries and the ORIGINAL Data with Pandas ---
    with open(boundaries_json_path, 'r') as f:
        boundaries = json.load(f)
    header_start = boundaries['header_start_index']
    data_end = boundaries['data_end_index']
    
    # Read the ORIGINAL file. Let pandas infer data types; it's designed to read formula values.
    # This is the most reliable way to get the numeric data.
    df = pd.read_excel(input_file, header=None, sheet_name=0)
    print("  [Read] Successfully loaded original Excel file into memory.")

    # --- Step 2: Slice and Process ---
    table_df = df.iloc[header_start : data_end + 1].copy()
    print(f"  [Slice] Extracted table from row {header_start} to {data_end}.")

    header_df = table_df.iloc[:HEADER_ROW_COUNT].copy()
    data_df = table_df.iloc[HEADER_ROW_COUNT:].copy()
    
    # Virtually unmerge headers using forward-fill
    header_df.ffill(axis=1, inplace=True)

    # Create correct and unique header names
    new_columns_raw = []
    for col_idx in range(header_df.shape[1]):
        levels = [str(header_df.iloc[row_idx, col_idx]) for row_idx in range(HEADER_ROW_COUNT)]
        cleaned_levels = [lvl.split('/')[0].strip() for lvl in levels if 'unnamed' not in lvl.lower() and lvl.lower() != 'nan']
        new_name = '_'.join(cleaned_levels).replace(' ', '_').replace('%_of', 'pct_of').lower()
        if new_name == 'department_department': new_name = 'department'
        new_columns_raw.append(new_name)
    
    final_columns = []
    counts = defaultdict(int)
    for name in new_columns_raw:
        counts[name] += 1
        if counts[name] > 1: final_columns.append(f"{name}_{counts[name]-1}")
        else: final_columns.append(name)
    print("  [Clean] Headers have been refactored and de-duplicated.")

    # Assign headers and clean the final DataFrame
    data_df.columns = final_columns
    first_col = data_df.columns[0]
    data_df.rename(columns={first_col: 'department'}, inplace=True)
    data_df = data_df[~data_df['department'].astype(str).str.contains('TOTAL|DEPARTMENTS', case=False, na=False)]
    data_df.dropna(how='all', inplace=True)
    data_df.reset_index(drop=True, inplace=True)

    # --- Step 3: Save Final Outputs ---
    data_df.to_excel(final_excel_path, index=False)
    print(f"  [Save] Final clean Excel file generated at '{final_excel_path}'")
    data_df.to_csv(final_csv_path, index=False)
    print(f"  [Save] Final clean CSV file generated at '{final_csv_path}'")