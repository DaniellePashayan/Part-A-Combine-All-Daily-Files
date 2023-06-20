import pandas as pd
import os
from glob import glob

username = os.getlogin()
folder_path = f'C:/Users/{username}/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - Hospital/CEMS/Daily Reports/'
sheet_name = "Detailed Report"


def combine_CEMS(month: str, year: str):
    if len(month) < 2:
        month = "0" + month
    # BusinessFile_Transaction Report - 2023-06-02 08-20-29 PM.xlsx
    try:
        files = pd.concat([pd.read_excel(file, sheet_name=sheet_name).assign(file_name=os.path.basename(file)) for file in glob(
            f"{folder_path}*{year}-{month}*.xlsx") if "Consolidated Files" not in file and "~" not in file])

        files.columns = files.columns.str.strip()  # remove leading and trailing spaces
        if not os.path.exists(f'{folder_path}/Consolidated Files/'):
            os.mkdir(f'{folder_path}/Consolidated Files/')
        files.to_excel(
            f'{folder_path}/Consolidated Files/{year} {month} Combined.xlsx', index=False)
        print(f"Files combined for {year} {month}")
    except ValueError as e:
        print(f"No files for {year} {month}: {e}")
