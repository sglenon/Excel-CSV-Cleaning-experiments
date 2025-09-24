# CSV/Excel Cleaning Pipeline

This project provides a robust pipeline for cleaning and processing Excel files with multiple sheets. It automates splitting, refreshing, boundary detection, and data cleaning for each sheet, producing ready-to-use cleaned Excel and CSV outputs.

## Features

- **Splits** each sheet of an Excel file into separate files
- **Refreshes** formulas and data connections (requires Microsoft Excel for full refresh)
- **Detects** table boundaries using AI (OpenAI API required)
- **Cleans** and standardizes data with pandas
- **Processes all sheets** in the input file, saving outputs with descriptive filenames

## Requirements

- Python 3.8+
- [uv](https://github.com/astral-sh/uv) (for fast script execution, or use `python` directly)
- Microsoft Excel (for formula/data refresh via `xlwings`)
- OpenAI API key (for table boundary detection)
- Dependencies listed in `requirements.txt`

## Installation

1. Clone the repository.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Set your OpenAI API key in a `.env` file:
   ```
   OPENAI_API_KEY=your-key-here
   ```

## Usage

Run the pipeline with:

```bash
uv run main.py <input_excel_file> <output_directory>
```
Or, using Python directly:
```bash
python main.py <input_excel_file> <output_directory>
```

- `<input_excel_file>`: Path to your source Excel file (e.g., `data/MyWorkbook.xlsx`)
- `<output_directory>`: Directory where all outputs will be saved (created if it doesn't exist)

## What the Pipeline Does

For **each sheet** in your Excel file, the pipeline will:

1. **Split** the sheet into its own Excel file.
2. **Refresh** formulas and data connections (in-place).
3. **Detect** the main data table boundaries using OpenAI GPT.
4. **Clean** and standardize the data with pandas.
5. **Save** cleaned Excel and CSV files, plus intermediate files, in the output directory.

### Output Files

For each sheet, the following files are generated in the output directory:

- `{basename}_sheet{idx}_{sheetname}.xlsx` — The split sheet file
- `{basename}_sheet{idx}_{sheetname}_boundaries.json` — Table boundary info (AI-generated)
- `{basename}_sheet{idx}_{sheetname}_cleaned.xlsx` — Cleaned Excel file
- `{basename}_sheet{idx}_{sheetname}_cleaned.csv` — Cleaned CSV file

*(`basename` is the original Excel filename without extension; `idx` is the sheet number; `sheetname` is the sanitized sheet name)*

## Example

```bash
uv run main.py data/MyWorkbook.xlsx results/
```

This will process all sheets in `MyWorkbook.xlsx` and save all outputs in the `results/` directory.

## Troubleshooting

- Ensure Microsoft Excel is installed and accessible for formula refresh.
- Set your OpenAI API key in `.env` for boundary detection.
- Check the console output for detailed logs and error messages.

## Project Structure

```
main.py                  # Orchestrates the full pipeline
src/
  sheets_to_excel.py     # Splits Excel into per-sheet files
  preprocessing_excel_sheets.py  # Refreshes formulas/data
  find_table_boundaries.py       # AI-based table boundary detection
  process_with_pandas.py         # Cleans and standardizes data
requirements.txt         # Python dependencies
```

## License

MIT License