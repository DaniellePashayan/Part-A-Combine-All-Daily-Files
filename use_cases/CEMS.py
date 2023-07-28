import pandas as pd
import os
from glob import glob
import sys
from tqdm import tqdm

import functions as fx

def combine(use_case_data: dict, month: str, year: str):

    if use_case_data["location"] == "SharePoint":
        daily_path = f'{fx.SHAREPOINT_PATH}/{use_case_data["daily_path"]}'
        consolidation_path = f'{fx.SHAREPOINT_PATH}/{use_case_data["consolidation_path"]}'
    else:
        daily_path = use_case_data["daily_path"]
        consolidation_path = use_case_data["consolidation_path"]

    sheet_name = use_case_data["sheet_name"]

    month = month.zfill(2)
    # BusinessFile_Transaction Report - 2023-06-02 08-20-29 PM.xlsx
    try:
        files = pd.concat([pd.read_excel(file, sheet_name=sheet_name).assign(file_name=os.path.basename(file)) for file in tqdm(glob(
            f"{daily_path}*{year}-{month}*.xlsx")) if "Consolidated Files" not in file and "~" not in file])

        files.columns = files.columns.str.strip()  # remove leading and trailing spaces
        if not os.path.exists(consolidation_path):
            os.mkdir(consolidation_path)
        files.to_excel(
            f'{consolidation_path}/{year} {month} Combined.xlsx', index=False)
    except ValueError as e:
        print(f"No files for {year} {month}: {e}")
