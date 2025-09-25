# script_b_process_with_pandas.py

import pandas as pd
import json
from collections import defaultdict

def process_table_with_pandas(input_file: str, boundaries_json_path: str, final_excel_path: str, final_csv_path: str):
    """
    Reads the original Excel file and uses the AI-found boundaries to perform
    a definitive, in-memory cleaning and structuring process with pandas.
    This version adaptively handles both simple and complex multi-level headers.
    """
    print("\n--- Step B: Processing Table with Pandas-First Approach ---")

    # --- Step 1: Load Boundaries and the ORIGINAL Data with Pandas ---
    with open(boundaries_json_path, 'r') as f:
        boundaries = json.load(f)
    header_start = boundaries['header_start_index']
    data_end = boundaries['data_end_index']
    
    df = pd.read_excel(input_file, header=None, sheet_name=0, dtype=str)
    print("  [Read] Successfully loaded original Excel file into memory as string data.")

    # --- Step 2: Slice and Process ---
    table_df = df.iloc[header_start : data_end + 1].copy().reset_index(drop=True)
    print(f"  [Slice] Extracted table from row {header_start} to {data_end}.")

    # --- Step 2a: ROBUST ADAPTIVE HEADER DETECTION ---
    # Heuristic: A "simple" header is a single row followed by a data row. A data row
    # typically has a value in the first column. A complex header has multiple header
    # rows, where the second row is often a sub-header and has no value in the first column.
    is_complex_header = True
    if len(table_df) > 1:
        # Check if the first cell of the SECOND row has data. If it does, it's a data row.
        second_row_first_cell = table_df.iloc[1, 0]
        if pd.notna(second_row_first_cell) and str(second_row_first_cell).strip():
            is_complex_header = False # It's a simple header because data starts on row 2.

    if not is_complex_header:
        # --- PATH A: SIMPLE HEADER ---
        print("  [Analyze] Detected a simple, single-row header. Using direct processing.")
        header_row_count = 1
        header_df = table_df.iloc[:header_row_count]
        data_df = table_df.iloc[header_row_count:].copy()
        
        new_columns_raw = [
            str(name).strip().replace(' ', '_').replace('%', 'pct').replace('/', '_').lower()
            for name in header_df.values[0]
        ]
        
    else:
        # --- PATH B: COMPLEX HEADER ---
        print("  [Analyze] Detected a complex, multi-row header. Applying dynamic analysis.")
        
        header_row_count = 1
        for i in range(1, min(5, len(table_df))):
            if pd.notna(table_df.iloc[i, 0]) and str(table_df.iloc[i, 0]).strip():
                header_row_count = i
                break
            header_row_count = i + 1
        print(f"  [Analyze] Dynamically determined header is {header_row_count} rows deep.")
        
        header_df = table_df.iloc[:header_row_count].copy()
        data_df = table_df.iloc[header_row_count:].copy()
        header_df.ffill(axis=1, inplace=True)

        new_columns_raw = []
        for col_idx in range(header_df.shape[1]):
            levels = [str(header_df.iloc[row_idx, col_idx]) for row_idx in range(header_row_count)]
            cleaned_levels = [lvl.strip() for lvl in levels if 'unnamed' not in lvl.lower() and lvl.lower() != 'nan']
            unique_levels = list(pd.Series(cleaned_levels).unique())
            new_name = '_'.join(unique_levels).replace(' ', '_').replace('%', 'pct').replace('/', '_').lower()
            new_columns_raw.append(new_name or f'unnamed_col_{col_idx}')

    # --- Step 2b: De-duplicate and Finalize Column Names ---
    final_columns = []
    counts = defaultdict(int)
    for name in new_columns_raw:
        counts[name] += 1
        final_columns.append(f"{name}_{counts[name]-1}" if counts[name] > 1 else name)
    print("  [Clean] Headers have been finalized and de-duplicated.")

    # --- Step 2c: Assign Headers and Clean Final DataFrame ---
    data_df.columns = final_columns
    
    first_column_name = data_df.columns[0]
    if first_column_name:
        data_df = data_df[~data_df[first_column_name].astype(str).str.contains('TOTAL|DEPARTMENTS', case=False, na=False)]
    
    data_df.dropna(how='all', inplace=True)
    data_df.reset_index(drop=True, inplace=True)

    # --- Step 3: Save Final Outputs ---
    data_df.to_excel(final_excel_path, index=False)
    print(f"  [Save] Final clean Excel file generated at '{final_excel_path}'")
    data_df.to_csv(final_csv_path, index=False)
    print(f"  [Save] Final clean CSV file generated at '{final_csv_path}'")