# main_script.py

import os
import sys
import argparse
from dotenv import load_dotenv
from pathlib import Path


# Pre-flight Check for API Key
load_dotenv()
if not os.getenv("OPENAI_API_KEY"):
    print("--- CRITICAL ERROR: OpenAI API key not found. Please set the OPENAI_API_KEY environment variable. ---")
    sys.exit(1)

import openai
openai.api_key = os.getenv("OPENAI_API_KEY")

try:
    from src.script_a_find_table_boundaries import find_table_boundaries
    from src.script_b_process_with_pandas import process_table_with_pandas
except ImportError as e:
    print(f"Error: Could not import necessary functions. Make sure all script files are in the same directory.")
    sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description="A robust pipeline to preprocess messy Excel files into clean CSVs.")
    parser.add_argument("input_excel_file", type=str, help="Path to the input messy Excel file.")
    parser.add_argument("output_directory", type=str, help="Path to the directory for the final output files.")
    
    args = parser.parse_args()
    input_path = Path(args.input_excel_file)
    output_dir = Path(args.output_directory)

    if not input_path.is_file():
        print(f"Error: Input file not found at '{input_path}'"); return

    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Define paths
    boundaries_json = output_dir / "table_boundaries.json"
    final_excel_path = output_dir / (input_path.stem + "_processed.xlsx")
    final_csv_path = output_dir / (input_path.stem + "_processed.csv")

    try:
        # Step A: Find the table coordinates
        find_table_boundaries(str(input_path), str(boundaries_json))
        
        # Step B: Use coordinates to process the original file directly
        process_table_with_pandas(
            input_file=str(input_path),
            boundaries_json_path=str(boundaries_json),
            final_excel_path=str(final_excel_path),
            final_csv_path=str(final_csv_path)
        )
        
        print("\n--- Pipeline Finished Successfully! ---")

    except Exception as e:
        print(f"\n--- An error occurred during the pipeline execution! ---")
        print(f"Error: {e}")

if __name__ == "__main__":
    main()