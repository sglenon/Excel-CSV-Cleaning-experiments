# script_a_find_table_boundaries.py

import pandas as pd
import openai
import json
import os

openai.api_key = os.getenv("OPENAI_API_KEY")

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