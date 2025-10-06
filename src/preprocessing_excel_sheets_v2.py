"""
Module for refreshing and extracting data from Excel sheets. (v2)

This module provides a function to programmatically open an Excel file,
refresh all data connections and formulas, save the file, and then load
the resulting data into a pandas DataFrame for further processing.

No major changes from v1, but included for version tracking and future extensibility.
"""

import pandas as pd
import openpyxl
import xlwings as xw  

import platform

def recalculate_and_refresh_sheets(input_file_path: str) -> pd.DataFrame:
    """
    Opens an Excel file, refreshes all data connections and formulas, saves the file,
    and loads the resulting data into a pandas DataFrame.

    On Linux, skips the refresh step and just loads the file with openpyxl.
    On Windows/macOS, uses xlwings to refresh formulas and data connections.

    Returns
    -------
    pd.DataFrame
        A DataFrame containing the values from the active worksheet of the refreshed Excel file.
    """
    system = platform.system()
    if system == "Linux":
        print("[WARN] Skipping Excel refresh step: xlwings/Excel not available on Linux.")
        wb_data = openpyxl.load_workbook(input_file_path, data_only=True)
        ws = wb_data.active
        df = pd.DataFrame(ws.values)
        return df
    else:
        # Start an invisible Excel application instance
        app_excel = xw.App(visible=False)
        # Open the specified Excel workbook
        wbk = app_excel.books.open(input_file_path)
        # Refresh all data connections and formulas in the workbook
        wbk.api.RefreshAll()
        # Save the workbook after refreshing
        wbk.save(input_file_path)
        # Close the workbook
        wbk.close()
        # Quit the Excel application
        app_excel.quit()

        # Load the workbook with openpyxl, reading only the values (not formulas)
        wb_data = openpyxl.load_workbook(input_file_path, data_only=True)
        ws = wb_data.active
        df = pd.DataFrame(ws.values)
        return df