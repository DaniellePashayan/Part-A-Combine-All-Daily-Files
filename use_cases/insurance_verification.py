import pandas as pd
import os
from glob import glob
import numpy as np
import re
import sys
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
    # BusinessFile_Transacti
    
    in_scope = ['MAGNACARE NW', 'MEDICARE', 'MEDICAID', 'AARP', 'NW DIRECT']

    def combine_files(month, year):
        files = pd.concat([pd.read_excel(file, sheet_name=sheet_name).assign(file_name=os.path.basename(file)) for file in glob(
            f"{daily_path}*{year}-{month}*.xlsx") if "Consolidated Files" not in file and "~" not in file])

        files.columns = files.columns.str.strip()  # remove leading and trailing spaces
        return files

    def check_in_scope_has_response(df):

        pattern = r'\b(?:{})'.format('|'.join(map(re.escape, in_scope)))
        # looks for the in scope insurances only if the response begins with that value, ie excludes Aetna Medicare but captures Medicare Inpatient
        scope_check = df[['Primary Insurance', 'Secondary Insurance', 'Tertiary Insurance']].apply(
            lambda x: x.str.contains(pattern, na=False))
        response_check = pd.DataFrame(index=df.index, columns=[
                                      'Primary Insurance', 'Secondary Insurance', 'Tertiary Insurance'])

        for col in ['Primary Insurance', 'Secondary Insurance', 'Tertiary Insurance']:
            response_check[col] = df[col +
                                     " Verification Status"] == "Verified"

        in_scope_and_response = (scope_check & response_check)
        df['In Scope with Response'] = in_scope_and_response.any(axis=1)
        return df

    def format(df):
        keep_cols = [
            'Account Number',
            'Visit Date',
            'Account Status',
            'Exception Type',
            'Bot Action',
            'Exception Reason',
            'Primary Insurance',
            'Secondary Insurance',
            'Tertiary Insurance',
            'Primary Insurance Status',
            'Secondary Insurance Status',
            'Tertiary Insurance Status',
            'Primary Insurance Verification Status',
            'Secondary Insurance Verification Status',
            'Tertiary Insurance Verification Status',
            'Facility',
            'Work List Name',
            'file_name',
            'In Scope with Response']
        df = df.loc[:, keep_cols]
        date_pattern = r'(\d{4}-\d{2}-\d{2})'
        df['file_name'] = df['file_name'].str.extract(date_pattern)
        df['file_name'] = pd.to_datetime(df['file_name']).dt.date
        df = df.rename({'file_name': 'File Date'})
        return df

    def classify_successes(df):
        output_conditions = [
            # Partially Processed with Insurance response for inscope payers
            (df['Bot Action'] == 'Partially Processed') & (
                df['In Scope with Response'] == True),
            # Partially Process with out Insurance response for inscope payers
            (df['Bot Action'] == 'Partially Processed') & (
                df['In Scope with Response'] == False),
            # Business Success
            (df['Exception Type'] == 'None'),
            (df['Exception Type'] == 'BusinessException'),
            (df['Exception Type'] == 'ApplicationException')
        ]

        output_choices = [
            'Partially Processed - Business Success',
            'Partially Processed - Business Exception',
            'Business Success',
            'Business Exception',
            'Application Exception'
        ]

        df['Output'] = np.select(
            output_conditions, output_choices, default=np.nan)
        return df

    df = combine_files(month, year)
    df = check_in_scope_has_response(df)
    df = format(df)
    df = classify_successes(df)
    try:
        df.to_excel(f'{consolidation_path}/{year} {month} Combined.xlsx', index=False)
        print(f"Files combined for {year} {month}")
    except ValueError as e:
        print(f"No files for {year} {month}: {e}")