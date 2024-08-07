import importlib
import json
import sys
import os
from datetime import datetime as dt
from loguru import logger

# Add the 'use_cases' directory to the sys.path
sys.path.append(os.path.join(os.path.dirname(__file__), 'use_cases'))

month = str(dt.now().strftime('%m')).zfill(2)
year = str(dt.now().strftime('%Y'))

def run_use_case(module_name, use_case_data, month, year):
    module = importlib.import_module(module_name)
    # Run the desired functionality for the module using use_case_data
    module.combine(use_case_data, month, year)

if __name__ == '__main__':
    # Read the JSON file
    with open('use_cases.json') as file:
        use_cases = json.load(file)

    # Iterate over the JSON and execute the modules
    for use_case_name, use_case_data in use_cases.items():
        module_name = use_case_data['module']
        print(f"Running {use_case_name}")
        run_use_case(module_name, use_case_data, month, year)