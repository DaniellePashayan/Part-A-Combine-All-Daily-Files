import pandas as pd
import os
from glob import glob
import sys
import functions as fx
from tqdm import tqdm

def combine(use_case_data: dict, month: str, year: str):

    if use_case_data["location"] == "SharePoint":
        daily_path = f'{fx.SHAREPOINT_PATH}/{use_case_data["daily_path"]}'
        consolidation_path = f'{fx.SHAREPOINT_PATH}/{use_case_data["consolidation_path"]}'
    else:
        daily_path = use_case_data["daily_path"]
        consolidation_path = use_case_data["consolidation_path"]

    sheet_names = use_case_data["sheet_name"]

    month = month.zfill(2)
    # BusinessFile_Transaction Report - 2023-06-02 08-20-29 PM.xlsx
    try:
        all_sheets = []
        for sheet_name in sheet_names:
            print(sheet_name)
            files = pd.concat([pd.read_excel(file, sheet_name=sheet_name).assign(file_name=os.path.basename(file)) for file in tqdm(glob( f"{daily_path}/PCx*{year}-{month}*.xlsx")) if "Consolidated Files" not in file and "~" not in file])
            all_sheets.append(files)
        all_sheets = pd.concat(all_sheets)

        all_sheets.columns = all_sheets.columns.str.strip()  # remove leading and trailing spaces
        
        os.makedirs(consolidation_path, exist_ok=True)
        all_sheets.to_excel(
            f'{consolidation_path}/{year} {month} PCx Combined.xlsx', index=False)
    except ValueError as e:
        print(f"No files for {year} {month}: {e}")

    
    return all_sheets