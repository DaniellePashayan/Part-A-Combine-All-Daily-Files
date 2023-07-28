import importlib
import json
import os
import sys

# Add the 'use_cases' directory to the sys.path
sys.path.append(os.path.join(os.path.dirname(__file__), 'use_cases'))

month = '07'
year = '2023'
use_case_name = 'SCM'

def run_use_case(module_name, use_case_data, month, year):
    module = importlib.import_module(module_name)
    # Run the desired functionality for the module using use_case_data
    module.combine(use_case_data, month, year)

# Read the JSON file
with open('use_cases.json') as file:
    use_cases = json.load(file)

# Iterate over the JSON and get the data for the use case name
use_case_data = use_cases[use_case_name]
module_name = use_case_data['module']
print(f"Running {use_case_name}")
run_use_case(module_name, use_case_data, month, year)
