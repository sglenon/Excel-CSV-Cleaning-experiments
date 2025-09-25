# src/process_with_pandas.py

import pandas as pd
import json
from collections import defaultdict

# --- GUARDRAIL CONFIGURATION ---
MIN_TABLE_ROWS = 3
MIN_TABLE_COLS = 2
MAX_HEADER_ROWS = 5
HEADER_TEXT_RATIO_THRESHOLD = 0.6
FOOTER_CLEANUP_KEYWORDS = ['TOTAL', 'GRAND TOTAL', 'SOURCE', '/1', 'NOTE'] # Simplified keywords for broader matching

def process_table_with_pandas(input_file: str, boundaries_json_path: str, final_excel_path: str, final_csv_path: str):
    """
    Reads the original Excel file and uses the AI-found boundaries to perform
    a definitive, in-memory cleaning and structuring process with pandas.
    This version includes guardrails to validate the AI's output and clean the final row.
    """
    print("\n--- Step B: Processing Table with Pandas-First Approach ---")

    # --- Step 1: Load Boundaries and the ORIGINAL Data with Pandas ---
    with open(boundaries_json_path, 'r') as f:
        boundaries = json.load(f)
    header_start = int(boundaries['header_start_index'])
    data_end = int(boundaries['data_end_index'])

    # GUARDRAIL 1: Logical Boundary Order
    if header_start >= data_end:
        raise ValueError(f"Guardrail Failed: The header_start_index ({header_start}) must be less than the data_end_index ({data_end}).")
    print("  [Guardrail] Passed: Boundary order is logical.")

    df = pd.read_excel(input_file, header=None, sheet_name=0, dtype=str)
    print("  [Read] Successfully loaded original Excel file into memory.")

    # --- Step 2: Slice and Process ---
    table_df = df.iloc[header_start : data_end + 1].copy().reset_index(drop=True)
    print(f"  [Slice] Extracted table from row {header_start} to {data_end}.")

    # GUARDRAIL 2: Minimum Table Dimensions
    rows, cols = table_df.shape
    if rows < MIN_TABLE_ROWS or cols < MIN_TABLE_COLS:
        raise ValueError(f"Guardrail Failed: The extracted table slice is too small ({rows}x{cols}).")
    print("  [Guardrail] Passed: Table dimensions are plausible.")

    # GUARDRAIL 3: Header Content Sanity Check
    header_candidate_row = table_df.iloc[0].dropna()
    if not header_candidate_row.empty:
        numeric_cells = pd.to_numeric(header_candidate_row, errors='coerce').notna().sum()
        text_ratio = (len(header_candidate_row) - numeric_cells) / len(header_candidate_row)
        if text_ratio < HEADER_TEXT_RATIO_THRESHOLD:
            raise ValueError(f"Guardrail Failed: Identified header at index {header_start} appears to be mostly numeric.")
    print("  [Guardrail] Passed: Header content appears to be text-based.")

    # --- Step 2a: Adaptive Header Detection ---
    is_complex_header = not (len(table_df) > 1 and pd.notna(table_df.iloc[1, 0]) and str(table_df.iloc[1, 0]).strip())
    
    if not is_complex_header:
        print("  [Analyze] Detected a simple, single-row header.")
        header_row_count = 1
        header_df, data_df = table_df.iloc[:1], table_df.iloc[1:].copy()
        new_columns_raw = [str(name).strip().replace(' ', '_').replace('%', 'pct').replace('/', '_').lower() for name in header_df.values[0]]
    else:
        print("  [Analyze] Detected a complex, multi-row header.")
        header_row_count = 1
        for i in range(1, min(MAX_HEADER_ROWS + 1, len(table_df))):
            if pd.notna(table_df.iloc[i, 0]) and str(table_df.iloc[i, 0]).strip():
                header_row_count = i; break
            header_row_count = i + 1

        # GUARDRAIL 4: Maximum Header Depth
        if header_row_count > MAX_HEADER_ROWS:
            raise ValueError(f"Guardrail Failed: Detected header depth ({header_row_count}) exceeds max of {MAX_HEADER_ROWS}.")
        print(f"  [Guardrail] Passed: Header depth ({header_row_count}) is within limits.")
        
        header_df, data_df = table_df.iloc[:header_row_count], table_df.iloc[header_row_count:].copy()
        header_df.ffill(axis=1, inplace=True)
        
        new_columns_raw = []
        for col_idx in range(header_df.shape[1]):
            levels = [str(header_df.iloc[r, col_idx]) for r in range(header_row_count)]
            cleaned_levels = [lvl.strip() for lvl in levels if 'unnamed' not in lvl.lower() and 'nan' not in lvl.lower()]
            new_name = '_'.join(list(pd.Series(cleaned_levels).unique())).replace(' ', '_').replace('%', 'pct').replace('/', '_').lower()
            new_columns_raw.append(new_name or f'unnamed_{col_idx}')

    # --- Step 2b: De-duplicate Columns ---
    final_columns = []
    counts = defaultdict(int)
    for name in new_columns_raw:
        counts[name] += 1
        final_columns.append(f"{name}_{counts[name]-1}" if counts[name] > 1 else name)
    print("  [Clean] Headers have been finalized and de-duplicated.")

    # --- Step 2c: Assign Headers and Final Cleanup ---
    data_df.columns = final_columns
    
    if not data_df.empty and data_df.columns[0]:
        first_col = data_df.columns[0]
        data_df = data_df[~data_df[first_col].astype(str).str.contains('TOTAL|DEPARTMENTS', case=False, na=False)]

    # --- NEW GUARDRAIL 5: Final Row Cleanup ---
    if not data_df.empty:
        last_row_first_cell = str(data_df.iloc[-1, 0]).strip().upper()
        if any(keyword in last_row_first_cell for keyword in FOOTER_CLEANUP_KEYWORDS):
            print(f"  [Guardrail] Detected summary/footer row: '{data_df.iloc[-1, 0]}'. Removing it.")
            data_df = data_df.iloc[:-1].copy()
    
    data_df.dropna(how='all', inplace=True)
    data_df.reset_index(drop=True, inplace=True)

    # --- Step 3: Save Final Outputs ---
    data_df.to_excel(final_excel_path, index=False)
    print(f"  [Save] Final clean Excel file generated at '{final_excel_path}'")
    data_df.to_csv(final_csv_path, index=False)
    print(f"  [Save] Final clean CSV file generated at '{final_csv_path}'")