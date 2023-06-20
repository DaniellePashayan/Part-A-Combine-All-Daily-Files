import pandas as pd
import os
from glob import glob
import pyodbc
import pytz

username = os.getlogin()
folder_path = f'C:/Users/{username}/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - Hospital/Lab/Daily Reports/'
sheet_name = "Detailed Report"
    
def dispathcer_combine(month, year):
    DB_SERVER = "SYKQTCOEURDB01V"
    DB_DATABASE = "UIPathInsights"

    username = os.getlogin()

    OUTPUT_PATH = f'C:/Users/{username}/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/Part A - Hospital/Lab/Queue Delay'

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

    if len(str(month)) == 1:
        month = f'0{month}'
    else:
        month = str(month)

    data.to_excel(f'{OUTPUT_PATH}/{str(year)} {month} Queue Delay.xlsx', index=None)
    print(f'File saved to {str(year)} {month}')

def combine_lab(month: str, year: str):
    if len(month) < 2:
        month = "0" + month
    #BusinessFile_Transaction Report - 2023-06-02 08-20-29 PM.xlsx
    try:
        files = pd.concat([pd.read_excel(file, sheet_name=sheet_name).assign(file_name=os.path.basename(file)) for file in glob(f"{folder_path}*{year}-{month}*.xlsx") if "Consolidated Files" not in file and "~" not in file])
        
        files.columns = files.columns.str.strip() # remove leading and trailing spaces
        
        if not os.path.exists(f'{folder_path}/Consolidated Files/'):
            os.mkdir(f'{folder_path}/Consolidated Files/')
        files.to_excel(f'{folder_path}/Consolidated Files/{year} {month} Combined.xlsx', index=False)
        print(f"Files combined for {year} {month}")
        dispathcer_combine(month, year)
        print(f"Dispatcher file created for {year} {month}")
    except ValueError as e:
        print(f"No files for {year} {month}: {e}")