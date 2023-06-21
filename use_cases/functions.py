import json
import os

def get_use_case_info(use_case_name):
    file = open('./use_cases.json', 'r')
    file = json.load(file)
    use_case_data = [use_case for use_case in file if use_case["name"] == use_case_name][0]
    return use_case_data

USER = os.getlogin()
SHAREPOINT_PATH = f'C:/Users/{USER}/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/'