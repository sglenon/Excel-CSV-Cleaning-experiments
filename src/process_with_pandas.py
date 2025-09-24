# script_b_process_with_pandas.py

import pandas as pd
import json
from collections import defaultdict

def process_table_with_pandas(input_file: str, boundaries_json_path: str, final_excel_path: str, final_csv_path: str):
    """
    Reads the original Excel file and uses the AI-found boundaries to perform
    a definitive, in-memory cleaning and structuring process with pandas.
    """
    print("\n--- Step B: Processing Table with Pandas-First Approach ---")

    # --- Step 1: Load Boundaries and the ORIGINAL Data with Pandas ---
    with open(boundaries_json_path, 'r') as f:
        boundaries = json.load(f)
    header_start = boundaries['header_start_index']
    data_end = boundaries['data_end_index']
    
    # Read all data as strings (`dtype=str`) to prevent any data loss,
    # formula misinterpretation, or automatic type conversion.
    df = pd.read_excel(input_file, header=None, sheet_name=0, dtype=str)
    print("  [Read] Successfully loaded original Excel file into memory as string data.")

    # --- Step 2: Slice and Process ---
    table_df = df.iloc[header_start : data_end + 1].copy().reset_index(drop=True)
    print(f"  [Slice] Extracted table from row {header_start} to {data_end}.")

    # --- Step 2a: Dynamically Find Header Row Count ---
    # This heuristic assumes that the first row of actual data will have a value
    # in the first column, whereas sub-header rows will not.
    header_row_count = 1
    # We scan up to the first 5 rows of the sliced table to find the end of the header
    for i in range(1, min(5, len(table_df))):
        first_cell_value = table_df.iloc[i, 0]
        if pd.notna(first_cell_value) and str(first_cell_value).strip():
            # This row has data in the first column, so we assume the header ended on the previous row.
            header_row_count = i
            break
        # If the first cell is empty, we assume it's another header row.
        header_row_count = i + 1
    
    print(f"  [Analyze] Dynamically determined header is {header_row_count} rows deep.")
    
    header_df = table_df.iloc[:header_row_count].copy()
    data_df = table_df.iloc[header_row_count:].copy()
    
    # Virtually unmerge headers using forward-fill
    header_df.ffill(axis=1, inplace=True)

    # --- Step 2b: Create Generalized and Unique Header Names ---
    new_columns_raw = []
    for col_idx in range(header_df.shape[1]):
        # Get all text levels for this column from the header rows
        levels = [str(header_df.iloc[row_idx, col_idx]) for row_idx in range(header_row_count)]
        
        # Clean the levels: remove junk values like 'nan' or 'Unnamed...'
        cleaned_levels = [lvl.strip() for lvl in levels if 'unnamed' not in lvl.lower() and lvl.lower() != 'nan']
        
        # Generalization: If a header and sub-header are identical, use the name only once.
        # This elegantly handles cases like 'DEPARTMENT', 'DEPARTMENT' -> 'department'
        unique_levels = list(pd.Series(cleaned_levels).unique())
        
        # Join the unique levels and apply SQL/RAG-friendly formatting
        new_name = '_'.join(unique_levels).replace(' ', '_').replace('%', 'pct').replace('/', '_').lower()
        
        # If after cleaning, the name is empty, provide a default name
        if not new_name:
            new_name = f'unnamed_col_{col_idx}'
        new_columns_raw.append(new_name)
    
    # De-duplicate any resulting column names (e.g., if two columns become 'value')
    final_columns = []
    counts = defaultdict(int)
    for name in new_columns_raw:
        counts[name] += 1
        if counts[name] > 1:
            final_columns.append(f"{name}_{counts[name]-1}")
        else:
            final_columns.append(name)
    print("  [Clean] Headers have been generalized and de-duplicated.")

    # --- Step 2c: Assign Headers and Clean Final DataFrame ---
    data_df.columns = final_columns
    
    # Rename the first column to a standard name if it exists
    if data_df.columns[0]:
        data_df.rename(columns={data_df.columns[0]: 'department'}, inplace=True)
        
    # Filter out total/summary rows and drop any fully empty rows
    data_df = data_df[~data_df['department'].astype(str).str.contains('TOTAL|DEPARTMENTS', case=False, na=False)]
    data_df.dropna(how='all', inplace=True)
    data_df.reset_index(drop=True, inplace=True)

    # --- Step 3: Save Final Outputs ---
    data_df.to_excel(final_excel_path, index=False)
    print(f"  [Save] Final clean Excel file generated at '{final_excel_path}'")
    data_df.to_csv(final_csv_path, index=False)
    print(f"  [Save] Final clean CSV file generated at '{final_csv_path}'")