# src/find_table_boundaries.py

import pandas as pd
import openai
import json
import os
import sys
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    print("--- CRITICAL ERROR in find_table_boundaries: OpenAI API key not found. ---")
    sys.exit(1)
openai.api_key = api_key

# --- CONFIGURATION FOR HEURISTICS ---
# How many rows to scan at the beginning of the file to find the header
HEADER_SEARCH_DEPTH = 100
# How many rows to sample from the end of the data to find the last data row
FOOTER_SEARCH_DEPTH = 100
# The minimum number of populated cells a row must have to be considered a potential header
MIN_CELLS_FOR_HEADER = 3

def find_candidate_header_region(df: pd.DataFrame) -> (pd.DataFrame, str):
    """
    Programmatically find a potential header region using heuristics.
    """
    print("  [Heuristic] Searching for candidate header region...")
    search_df = df.head(HEADER_SEARCH_DEPTH).copy()
    # Count non-NA cells per row
    search_df['non_na_count'] = search_df.notna().sum(axis=1)
    
    # Find the first row that meets our criteria for a potential header
    potential_header_indices = search_df[search_df['non_na_count'] >= MIN_CELLS_FOR_HEADER].index
    
    if potential_header_indices.empty:
        print("  [Warning] No dense rows found to indicate a header. Using first 20 rows as candidate.")
        start_index = 0
    else:
        # We assume the header starts around the first dense row found
        start_index = max(0, potential_header_indices[0] - 5) # Go 5 rows back for context
        
    end_index = start_index + 20 # Provide a 20-row window for the AI
    
    candidate_df = df.iloc[start_index:end_index]
    print(f"  [Heuristic] Identified candidate header region from index {start_index} to {end_index-1}.")
    return candidate_df, candidate_df.to_string(index=True, header=False)

def find_candidate_footer_region(df: pd.DataFrame) -> (pd.DataFrame, str):
    """
    Programmatically find a potential footer region near the end of the data.
    """
    print("  [Heuristic] Searching for candidate footer region...")
    # Find the last row with any data
    last_non_empty_row_index = df.dropna(how='all').index.max()
    
    if pd.isna(last_non_empty_row_index):
        print("  [Warning] No data found in file. Using last 100 rows.")
        start_index = max(0, len(df) - FOOTER_SEARCH_DEPTH)
    else:
        start_index = max(0, last_non_empty_row_index - FOOTER_SEARCH_DEPTH)

    candidate_df = df.iloc[start_index : last_non_empty_row_index + 1]
    print(f"  [Heuristic] Identified candidate footer region from index {start_index} to {last_non_empty_row_index}.")
    return candidate_df, candidate_df.to_string(index=True, header=False)


# In src/find_table_boundaries.py

def get_boundary_from_ai(text_chunk: str, task: str) -> int:
    """
    Makes a focused API call to OpenAI to get a single boundary index.
    """
    if task == "header":
        prompt_text = "Analyze this excerpt from the top of an Excel sheet. Identify the row index for the primary header row. This is the row containing the main column titles, located *immediately above* the first row of actual data. The index is the number on the far left. Respond with ONLY a JSON object containing the key 'header_start_index'."
        key = "header_start_index"
    elif task == "footer":
        # --- REFINED PROMPT ---
        prompt_text = (
            "You are a meticulous data analyst. Analyze this excerpt from the bottom of a data table. "
            "Your task is to identify the row index for the **final row of actual data**.\n\n"
            "CRITICAL: **DO NOT** include the index of any summary row (e.g., 'TOTAL', 'Grand Total') or any footnote row (e.g., 'Source:', '/1 Notes'). "
            "The index you return must be the last row containing regular data entries.\n\n"
            "The index is the number on the far left. Respond with ONLY a JSON object containing the key 'data_end_index'."
        )
        key = "data_end_index"
    else:
        raise ValueError("Task must be 'header' or 'footer'")

    # ... (rest of the function is the same)
    response = openai.chat.completions.create(
        model="gpt-4-turbo",
        messages=[{"role": "system", "content": prompt_text}, {"role": "user", "content": text_chunk}],
        response_format={"type": "json_object"}
    )
    
    content = json.loads(response.choices[0].message.content)
    if key not in content:
        raise ValueError(f"AI response did not contain the required key: {key}")
    
    return int(content[key])


def find_table_boundaries(file_path: str, output_json_path: str):
    """
    Uses a hybrid heuristic-AI approach to find precise table boundaries in large files.
    """
    print("--- Step A: Finding Table Boundaries (Hybrid Approach) ---")
    try:
        df = pd.read_excel(file_path, header=None, sheet_name=0, dtype=str)

        # 1. Find Header
        _, header_text = find_candidate_header_region(df)
        header_start_index = get_boundary_from_ai(header_text, "header")
        print(f"  [AI] Confirmed header start: {header_start_index}")

        # 2. Find Footer
        _, footer_text = find_candidate_footer_region(df)
        data_end_index = get_boundary_from_ai(footer_text, "footer")
        print(f"  [AI] Confirmed data end: {data_end_index}")

        # 3. Validation
        if header_start_index >= data_end_index:
            raise ValueError(f"Validation failed: header_start_index ({header_start_index}) cannot be after data_end_index ({data_end_index}).")

        boundaries = {
            "header_start_index": header_start_index,
            "data_end_index": data_end_index
        }
        
        with open(output_json_path, 'w') as f:
            json.dump(boundaries, f, indent=4)
        print(f"  [AI] Table boundaries saved to '{output_json_path}'")

    except Exception as e:
        print(f"  [Error] An error occurred in find_table_boundaries: {e}")
        raise