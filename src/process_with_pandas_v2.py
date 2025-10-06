# script_b_process_with_pandas_v2.py

import pandas as pd
import json
import os
import re
from collections import defaultdict

def clean_header_name(name):
    # Remove nonsense characters: / _ ? ! @ # % & * ( ) [ ] { } ; : , . < > | \ " ' and spaces
    cleaned = re.sub(r'[\/_\?\!\@\#\%\&\*\(\)\[\]\{\}\;\:\,\.\<\>\|\\\"\']', '', str(name))
    cleaned = cleaned.replace(' ', '_')
    cleaned = cleaned.lower()
    return cleaned

def extract_commentaries(df, header_start, data_end):
    # Extracts commentaries: any cell outside main data table or with footnote marker (e.g., /1, *, etc.)
    commentaries = []
    footnote_pattern = re.compile(r'(^|\s)[\/\*][0-9a-zA-Z]+')
    for idx, row in df.iterrows():
        if idx < header_start or idx > data_end:
            for col, val in enumerate(row):
                if pd.notna(val) and str(val).strip():
                    commentaries.append({'row': int(idx), 'col': int(col), 'value': str(val)})
        else:
            for col, val in enumerate(row):
                if pd.notna(val) and str(val).strip() and footnote_pattern.search(str(val)):
                    commentaries.append({'row': int(idx), 'col': int(col), 'value': str(val)})
    return commentaries

def process_table_with_pandas(input_file: str, boundaries_json_path: str, final_excel_path: str, final_csv_path: str, metadata_json_path: str = None):
    """
    Reads the original Excel file and uses the AI-found boundaries to perform
    a definitive, in-memory cleaning and structuring process with pandas.
    This version adaptively handles both simple and complex multi-level headers.
    Also extracts commentaries as metadata and saves to JSON if path is provided.
    """
    print("\n--- Step B: Processing Table with Pandas-First Approach (v2) ---")

    # --- Step 1: Load Boundaries and the ORIGINAL Data with Pandas ---
    with open(boundaries_json_path, 'r') as f:
        boundaries = json.load(f)
    header_start = boundaries['header_start_index']
    data_end = boundaries['data_end_index']
    
    df = pd.read_excel(input_file, header=None, sheet_name=0, dtype=str)
    print("  [Read] Successfully loaded original Excel file into memory as string data.")

    # --- Step 2: Extract and Save Metadata (Commentaries) ---
    commentaries = extract_commentaries(df, header_start, data_end)
    if metadata_json_path:
        os.makedirs(os.path.dirname(metadata_json_path), exist_ok=True)
        with open(metadata_json_path, 'w') as f:
            json.dump(commentaries, f, indent=2)
        print(f"  [Meta] Extracted {len(commentaries)} commentaries to '{metadata_json_path}'")

    # --- Step 3: Slice and Process ---
    table_df = df.iloc[header_start : data_end + 1].copy().reset_index(drop=True)
    print(f"  [Slice] Extracted table from row {header_start} to {data_end}.")

    # --- Step 3a: ROBUST ADAPTIVE HEADER DETECTION ---
    is_complex_header = True
    if len(table_df) > 1:
        second_row_first_cell = table_df.iloc[1, 0]
        if pd.notna(second_row_first_cell) and str(second_row_first_cell).strip():
            is_complex_header = False

    if not is_complex_header:
        # --- PATH A: SIMPLE HEADER ---
        print("  [Analyze] Detected a simple, single-row header. Using direct processing.")
        header_row_count = 1
        header_df = table_df.iloc[:header_row_count]
        data_df = table_df.iloc[header_row_count:].copy()
        
        new_columns_raw = [
            clean_header_name(name)
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
            cleaned_levels = [clean_header_name(lvl) for lvl in levels if 'unnamed' not in lvl.lower() and lvl.lower() != 'nan']
            unique_levels = list(pd.Series(cleaned_levels).unique())
            new_name = '_'.join(unique_levels)
            new_columns_raw.append(new_name or f'unnamed_col_{col_idx}')

    # --- Step 3b: De-duplicate and Finalize Column Names ---
    final_columns = []
    counts = defaultdict(int)
    for name in new_columns_raw:
        counts[name] += 1
        final_columns.append(f"{name}_{counts[name]-1}" if counts[name] > 1 else name)
    print("  [Clean] Headers have been finalized and de-duplicated.")

    # --- Step 3c: Assign Headers and Clean Final DataFrame ---
    data_df.columns = final_columns
    
    first_column_name = data_df.columns[0]
    if first_column_name:
        data_df = data_df[~data_df[first_column_name].astype(str).str.contains('TOTAL|DEPARTMENTS', case=False, na=False)]
    
    data_df.dropna(how='all', inplace=True)
    data_df.reset_index(drop=True, inplace=True)

    # --- Step 4: Save Final Outputs ---
    data_df.to_excel(final_excel_path, index=False)
    print(f"  [Save] Final clean Excel file generated at '{final_excel_path}'")
    data_df.to_csv(final_csv_path, index=False)
    print(f"  [Save] Final clean CSV file generated at '{final_csv_path}'")