# script_a_find_table_boundaries_v2.py

import pandas as pd
import openai
import json
import os
import sys
from dotenv import load_dotenv

# --- Load environment variables from .env file ---
load_dotenv()

# --- Add a robust pre-flight check for the API key ---
api_key = os.getenv("OPENAI_API_KEY")
if not api_key:
    print("--- CRITICAL ERROR in Script A: OpenAI API key not found. ---")
    print("--- Please ensure your .env file exists and contains the OPENAI_API_KEY. ---")
    sys.exit(1)
openai.api_key = api_key

SAMPLE_ROW_COUNT = 40

def find_table_boundaries(file_path: str, output_json_path: str):
    """
    Uses pandas to read the original file and AI to find the precise table boundaries.
    Samples large files to avoid token limits and uses a robust prompt.
    """
    print("--- Step A: Finding Table Boundaries using Pandas (v2) ---")
    try:
        df = pd.read_excel(file_path, header=None, sheet_name=0, dtype=str)

        if len(df) > (SAMPLE_ROW_COUNT * 2):
            print(f"  [Sample] File is large. Creating a sample of the first and last {SAMPLE_ROW_COUNT} rows.")
            head_df = df.head(SAMPLE_ROW_COUNT)
            tail_df = df.tail(SAMPLE_ROW_COUNT)
            df_string = (
                head_df.to_string(index=True, header=False) +
                "\n\n [... OMITTED MIDDLE ROWS ...] \n\n" +
                tail_df.to_string(index=True, header=False)
            )
        else:
            print("  [Sample] File is small. Using the full content.")
            df_string = df.to_string(index=True, header=False)

        prompt = (
            "You are a meticulous data analyst. Your task is to find the exact boundaries of the main data table in the provided text from an Excel sheet.\n\n"
            "1.  **header_start_index**: Find the row index for the primary header row. This is the row containing the main column titles, located *immediately above* the first row of actual data. The index is the number on the far left.\n"
            "2.  **data_end_index**: Find the row index for the final row of data. This is the last entry before any summary totals or footnotes (e.g., /1 Source...).\n\n"
            "Analyze carefully. Respond with ONLY a JSON object. For example: {header_start_index: 3, data_end_index: 52}"
        )

        response = openai.chat.completions.create(
            model="gpt-4-turbo",
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