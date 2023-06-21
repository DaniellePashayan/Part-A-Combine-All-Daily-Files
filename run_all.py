import importlib
import json
import sys
import os

# Add the 'use_cases' directory to the sys.path
sys.path.append(os.path.join(os.path.dirname(__file__), 'use_cases'))

month = '06'
year = '2023'


def run_use_case(module_name, use_case_data, month, year):
    module = importlib.import_module(module_name)
    # Run the desired functionality for the module using use_case_data
    module.combine(use_case_data, month, year)


# Read the JSON file
with open('use_cases.json') as file:
    use_cases = json.load(file)

# Iterate over the JSON and execute the modules
for use_case_name, use_case_data in use_cases.items():
    module_name = use_case_data['module']
    print(f"Running {use_case_name}")
    run_use_case(module_name, use_case_data, month, year)
