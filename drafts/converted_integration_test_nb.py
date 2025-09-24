# %%
import os
import sys
import argparse
from dotenv import load_dotenv
from pathlib import Path
import pandas as pd
import openai
import tempfile
import shutil
import re
from pathlib import Path
import openpyxl
import xlwings as xw

# %% [markdown]
# # Loading of Data-Only Excel Files

# %% [markdown]
# ### Defining Functions

# %%
app_excel = xw.App(visible=False) 
input_file = r"C:\Users\CSD Admin\OneDrive - DOST-ASTI\Kevin\CODING\CSV_Excel_Cleaning\test_sheet_by_department.xlsx"
wbk = app_excel.books.open(input_file)
wbk.api.RefreshAll()
wbk.save(input_file)
wbk.close()
app_excel.quit()

# %%
wb_data = openpyxl.load_workbook(input_file, data_only=True)
ws = wb_data["By Department"]
df = pd.DataFrame(ws.values)
df.head(100)

# %%
#with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
#    temp_path = temp_file.name
temp_path = r"C:\Users\CSD Admin\OneDrive - DOST-ASTI\Kevin\CODING\CSV_Excel_Cleaning\results\temp_file.xlsx"
input_1 = input_file
output_evaluated = evaluate_formulas_in_excel(input_1, temp_path)
df = pd.read_excel(temp_path)
df.head(100)

# %% [markdown]
# # Defining Table Boundaries 

# %% [markdown]
# ## Loading Environment

# %%
load_dotenv()
if not os.getenv("OPENAI_API_KEY"):
    print("--- CRITICAL ERROR: OpenAI API key not found. Please set the OPENAI_API_KEY environment variable. ---")
    sys.exit(1)

import openai
openai.api_key = os.getenv("OPENAI_API_KEY")

# %%
try:
    from src.script_a_find_table_boundaries import find_table_boundaries
    from src.script_b_process_with_pandas import process_table_with_pandas
except ImportError as e:
    print(f"Error: Could not import necessary functions. Make sure all script files are in the same directory.")
    sys.exit(1)

# %% [markdown]
# ## find_table_boundaries function

# %%
def find_table_boundaries(file_path: str, output_json_path: str):
    """
    Uses pandas to read the original file and AI to find the precise table boundaries.
    """
    print("--- Step A: Finding Table Boundaries using Pandas ---")
    try:
        # Read the file with pandas, forcing everything to string for the AI analysis text.
        # This prevents any data loss during this initial inspection step.
        df = pd.read_excel(file_path, header=None, sheet_name=0, dtype=str)
        df_string = df.to_string(index=True, header=False)

        prompt = (
            "You are an expert data analyst. Below is text from an Excel sheet. "
            "Your task is to find the boundaries of the main data table.\n\n"
            "1.  **header_start_index**: Find the row index where the main column headers begin. This row contains 'DEPARTMENT', 'NCA RELEASES', etc. The index is the number on the far left.\n"
            "2.  **data_end_index**: Find the row index of the last row of actual data (the last department/agency). This is the row just before the footnotes (which start with '/1 Source...').\n\n"
            "Respond with ONLY a JSON object containing these two keys. For example: {\"header_start_index\": 4, \"data_end_index\": 52}"
        )

        response = openai.chat.completions.create(
            model="gpt-4-turbo-preview",
            messages=[{"role": "system", "content": prompt}, {"role": "user", "content": df_string}],
            response_format={"type": "json_object"}
        )
        
        boundaries = json.loads(response.choices[0].message.content)
        
        if 'header_start_index' not in boundaries or 'data_end_index' not in boundaries:
            raise ValueError("AI response did not contain the required keys.")

        print(f"  [AI] Identified header start: {boundaries['header_start_index']}, data end: {boundaries['data_end_index']}")
        with open(output_json_path, 'w') as f:
            json.dump(boundaries, f, indent=4)
        print(f"  [AI] Table boundaries saved to '{output_json_path}'")

    except Exception as e:
        print(f"  [Error] An error occurred in Script A: {e}")
        raise

# %% [markdown]
# ## Running function on input file

# %%
import json 
input_excel_file = r"C:\Users\CSD Admin\OneDrive - DOST-ASTI\Kevin\CODING\CSV_Excel_Cleaning\test_sheet_by_department.xlsx"
output_directory = r"C:\Users\CSD Admin\OneDrive - DOST-ASTI\Kevin\CODING\CSV_Excel_Cleaning\results"

input_path = Path(input_excel_file)
output_dir = Path(output_directory)

if not input_path.is_file():
    raise FileNotFoundError(f"Error: Input file not found at '{input_path}'")

output_dir.mkdir(parents=True, exist_ok=True)

# Define paths
boundaries_json = output_dir / "table_boundaries.json"
final_excel_path = output_dir / (input_path.stem + "_processed.xlsx")
final_csv_path = output_dir / (input_path.stem + "_processed.csv")

# to run, uv run main_script.py test.xlsx output_folder

step_a_results = find_table_boundaries(str(input_path), str(boundaries_json))

# %% [markdown]
# # Processing with Pandas

# %%
from collections import defaultdict
HEADER_ROW_COUNT = 2

# %%
# --- Step 1: Load Boundaries and the ORIGINAL Data with Pandas ---
boundaries_json_path = r"C:\Users\CSD Admin\OneDrive - DOST-ASTI\Kevin\CODING\CSV_Excel_Cleaning\results\table_boundaries.json"

with open(boundaries_json_path, 'r') as f:
    boundaries = json.load(f)
header_start = boundaries['header_start_index']
data_end = boundaries['data_end_index']

# Read the ORIGINAL file. Let pandas infer data types; it's designed to read formula values.
# This is the most reliable way to get the numeric data.
df = pd.read_excel(input_path, header=None, sheet_name=0)
df.head(100)

# %%
print("  [Read] Successfully loaded original Excel file into memory.")
# --- Step 2: Slice and Process ---
table_df = df.iloc[header_start : data_end + 1].copy()
print(f"  [Slice] Extracted table from row {header_start} to {data_end}.")
header_df = table_df.iloc[:HEADER_ROW_COUNT].copy()
data_df = table_df.iloc[HEADER_ROW_COUNT:].copy()

# %%
header_df.ffill(axis=1, inplace=True)

# %%
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
print(final_columns)

# %%
data_df.columns = final_columns
first_col = data_df.columns[0]
data_df.head(100)

# %%
data_df.rename(columns={first_col: 'department'}, inplace=True)


# %%
data_df = data_df[~data_df['department'].astype(str).str.contains('TOTAL|DEPARTMENTS', case=False, na=False)]
data_df.head()

# %%

data_df.dropna(how='all', inplace=True)
data_df.reset_index(drop=True, inplace=True)
data_df.head(100)

# %%
data_df.to_excel(final_excel_path, index=False)
print(f"  [Save] Final clean Excel file generated at '{final_excel_path}'")
data_df.to_csv(final_csv_path, index=False)
print(f"  [Save] Final clean CSV file generated at '{final_csv_path}'")


