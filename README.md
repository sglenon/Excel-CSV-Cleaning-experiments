# Excel Preprocessing Pipeline

This project provides a robust, two-step pipeline for cleaning and preprocessing messy Excel files. It intelligently identifies the boundaries of the main data table within a sheet using an AI model (GPT-4) and then processes that table with the powerful `pandas` library to produce clean, machine-readable CSV and Excel files.

## How It Works

The pipeline is orchestrated by `main_script.py` and broken into two distinct stages:

### Step A: AI-Powered Table Boundary Detection (`script_a_find_table_boundaries.py`)

1.  **Read as Text:** The script first reads the entire content of the specified Excel sheet into a `pandas` DataFrame, intentionally converting all cells to strings. This ensures the raw text, including headers and notes, is preserved for analysis.
2.  **AI Analysis:** The string representation of the sheet is sent to the OpenAI GPT-4 API with a specific prompt. The prompt instructs the AI to act as a data analyst and identify the starting row index of the main table's headers and the ending row index of the data, just before any footers or totals.
3.  **Save Coordinates:** The AI returns a JSON object containing these two coordinates (`header_start_index` and `data_end_index`), which is then saved to the output directory (e.g., `results/table_boundaries.json`).

### Step B: Pandas-First Data Processing (`script_b_process_with_pandas.py`)

1.  **Load Coordinates:** This script loads the coordinates identified by the AI in Step A.
2.  **Precise Read:** It re-reads the *original* Excel file into a `pandas` DataFrame. This time, it allows `pandas` to infer data types, correctly interpreting numbers, dates, and formulas without any data loss.
3.  **Slice and Dice:** Using the AI-provided coordinates, the script slices the DataFrame to isolate the exact table, separating the multi-level headers from the data rows.
4.  **Header Cleaning:** It programmatically "un-merges" the headers by forward-filling values, cleans up the text, and combines the multi-level rows into a single, unique, and clean header row (e.g., `NCA RELEASES` and `% of Total` become `nca_releases_pct_of_total`).
5.  **Data Cleanup:** The script assigns the new, clean headers to the DataFrame, filters out any summary rows (like 'TOTAL'), and resets the index for a clean final output.
6.  **Save Outputs:** The final, cleaned DataFrame is saved as both a new Excel file (`*_processed.xlsx`) and a CSV file (`*_processed.csv`).

## Problem
During the process, there is a loss of data.
While there are now headers and no commentaries, there are some columns WITH NO DATA. I think there is a mistake in the step to process the original file which I do not know how to debug.

## Prerequisites

*   Python 3.7+
*   An OpenAI API Key
*   Required Python libraries: `pandas`, `openai`

## Setup

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd <repository-directory>
    ```

2.  **Install dependencies:**
    ```bash
    pip install pandas openai openpyxl
    ```

3.  **Set your OpenAI API Key:**
    The script requires the `OPENAI_API_KEY` to be set as an environment variable.

    *   **macOS/Linux:**
        ```bash
        export OPENAI_API_KEY='your_api_key_here'
        ```
    *   **Windows (Command Prompt):**
        ```bash
        set OPENAI_API_KEY=your_api_key_here
        ```
    *   **Windows (PowerShell):**
        ```bash
        $env:OPENAI_API_KEY="your_api_key_here"
        ```

## Usage

Run the main script from your terminal, providing the path to the messy input Excel file and the desired output directory.

```bash
python main_script.py <path/to/your/input_file.xlsx> <path/to/your/output_directory>
```

### Example

```bash
python 06/main_script.py path/to/messy_data.xlsx 06/results
```

Upon successful execution, the `06/results` directory will contain:
*   `table_boundaries.json`: The coordinates identified by the AI.
*   `messy_data_processed.xlsx`: The final, cleaned data in Excel format.
*   `messy_data_processed.csv`: The final, cleaned data in CSV format.

## File Structure

```
└── 06/
    ├── main_script.py                  # Main pipeline orchestrator
    ├── script_a_find_table_boundaries.py # Step A: AI boundary detection
    ├── script_b_process_with_pandas.py # Step B: Pandas data cleaning
    └── results/                          # Default output directory
        └── table_boundaries.json         # Example AI output
```
