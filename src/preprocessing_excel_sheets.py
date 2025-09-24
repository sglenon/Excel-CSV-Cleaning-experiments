"""
Module for refreshing and extracting data from Excel sheets.

This module provides a function to programmatically open an Excel file,
refresh all data connections and formulas, save the file, and then load
the resulting data into a pandas DataFrame for further processing.
"""

import pandas as pd
import openpyxl
import xlwings as xw  

def recalculate_and_refresh_sheets(input_file_path: str) -> pd.DataFrame:
    """
    Opens an Excel file, refreshes all data connections and formulas, saves the file,
    and loads the resulting data into a pandas DataFrame.

    This function uses xlwings to automate Excel, ensuring that all formulas and
    external data connections are recalculated and up to date. After saving the
    refreshed file, it uses openpyxl to read the values from the active worksheet
    and returns them as a pandas DataFrame.

    Parameters
    ----------
    input_file_path : str
        The file path to the Excel workbook to be refreshed and read.

    Returns
    -------
    pd.DataFrame
        A DataFrame containing the values from the active worksheet of the refreshed Excel file.

    Notes
    -----
    - Requires Microsoft Excel to be installed on the system (for xlwings).
    - Only the active worksheet is loaded into the DataFrame.
    - Formulas are evaluated and only their resulting values are returned.
    """


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
    # Get the active worksheet from the workbook
    ws = wb_data.active
    # Convert the worksheet values to a pandas DataFrame
    df = pd.DataFrame(ws.values)
    return df