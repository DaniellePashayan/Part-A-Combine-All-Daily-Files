import pandas as pd
import os
from glob import glob
import pyodbc
import pytz
import functions as fx
from tqdm import tqdm

def combine(use_case_data: dict, month: str, year: str):

    if use_case_data["location"] == "SharePoint":
        daily_path = f'{fx.SHAREPOINT_PATH}/{use_case_data["daily_path"]}'
        consolidation_path = f'{fx.SHAREPOINT_PATH}/{use_case_data["consolidation_path"]}'
    else:
        daily_path = use_case_data["daily_path"]
        consolidation_path = use_case_data["consolidation_path"]

    sheet_name = use_case_data["sheet_name"]

    month = month.zfill(2)
    
    def dispathcer_combine(month, year):
        DB_SERVER = "SYKQTCOEURDB01V"
        DB_DATABASE = "UIPathInsights"

        username = os.getlogin()

        OUTPUT_PATH = f'C:/Users/{fx.USER}/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - Hospital/Lab/Queue Delay'

        cnxn = pyodbc.connect(
            driver="SQL Server Native Client 11.0",
            server=DB_SERVER,
            Database=DB_DATABASE,
            Trusted_Connection="yes"
        )
        cursor = cnxn.cursor()

        # Build up our query string
        query = f"""
            SELECT
                converted.CreationTime_EST,
                converted.StartProcessing_EST,
                converted.QueueName,
                converted.RobotName,
                converted.[Accession Number],
                converted.[Payor Selection]
            FROM
                (
                    SELECT
                        CONVERT(DATETIME, SWITCHOFFSET(QI.CreationTime, DATEPART(TZOFFSET, QI.CreationTime AT TIME ZONE 'UTC' AT TIME ZONE 'Eastern Standard Time'))) AS 'CreationTime_EST',
                        CONVERT(DATETIME, SWITCHOFFSET(QI.StartProcessing, DATEPART(TZOFFSET, QI.StartProcessing AT TIME ZONE 'UTC' AT TIME ZONE 'Eastern Standard Time'))) AS 'StartProcessing_EST',
                        QI.QueueName,
                        QI.RobotName,
                        JSON_VALUE(QI.SpecificData, '$.DynamicProperties.Accession') AS 'Accession Number',
                        JSON_VALUE(QI.SpecificData, '$.DynamicProperties."Payor Selection"') AS 'Payor Selection'
                    FROM
                        dbo.QueueItems QI
                    WHERE
                        QI.QueueName = 'LAB_Queue'
                        AND QI.ProcessingStatus = 3
                        AND YEAR(QI.CreationTime) = {year}
                        AND MONTH(QI.CreationTime) = {month}
                ) AS converted
            ORDER BY
                converted.CreationTime_EST DESC
        """

        data = pd.read_sql(query, cnxn)

        # # number of days between startprocessing est and creationtime est
        data['ProcessingDelay'] = (data['StartProcessing_EST'] - data['CreationTime_EST']).dt.days

        data.to_excel(f'{OUTPUT_PATH}/{str(year)} {month} Queue Delay.xlsx', index=None)
        print(f'File saved to {str(year)} {month}')

    def combine_output(month: str, year: str):
        #BusinessFile_Transaction Report - 2023-06-02 08-20-29 PM.xlsx
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
    
    dispathcer_combine(month, year)
    combine_output(month, year)