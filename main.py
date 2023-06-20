import json
import os
import tkinter as tk
import pandas as pd
from glob import glob
from tkinter import messagebox

from use_cases.CEMS import combine_CEMS
from use_cases.CHHA import combine_CHHA
from use_cases.claim_status import combine_claim_status
from use_cases.clinical_appeals import combine_clinical_appeals
from use_cases.clinical_denials import combine_clinical_denials
from use_cases.insurance_verification import combine_IV
from use_cases.lab import combine_lab
from use_cases.minireg import combine_minireg
from use_cases.prereg import combine_prereg


user = os.getlogin()
sharepoint_path = f'C:/Users/{user}/Northwell Health/CBO (1111 Marcus Ave M04) - Robotic Process Automation/'


def add_new_use_case():
    # Function to handle the "add new use case" button click event
    new_window = tk.Toplevel(root)
    new_window.title("Add New Use Case")
    new_window.geometry("600x400")

    # Use Case Name
    tk.Label(new_window, text="Use Case Name:").pack()
    use_case_name_entry = tk.Entry(new_window)
    use_case_name_entry.pack()

    # Daily Path
    tk.Label(new_window, text="Daily Path:").pack()
    daily_path_entry = tk.Entry(new_window)
    daily_path_entry.pack()

    # Consolidation Path
    tk.Label(new_window, text="Consolidation Path:").pack()
    consolidation_path_entry = tk.Entry(new_window)
    consolidation_path_entry.pack()

    # Sheet Name
    tk.Label(new_window, text="Sheet Name:").pack()
    sheet_name_entry = tk.Entry(new_window)
    sheet_name_entry.pack()

    def save_use_case():
        # Function to handle saving the new use case information
        use_case_name = use_case_name_entry.get()
        daily_path = daily_path_entry.get()
        consolidation_path = consolidation_path_entry.get()

        # Add the new use case information to the existing JSON file
        with open("./use_cases.json", "r+") as file:
            data = json.load(file)
            new_use_case = {
                "name": use_case_name,
                "daily_path": daily_path,
                "consolidation_path": consolidation_path
            }
            data.append(new_use_case)
            file.seek(0)
            json.dump(data, file, indent=4)
            file.truncate()

        messagebox.showinfo("Success", "New use case added successfully.")
        new_window.destroy()
        render_checkboxes()

    # Save Button
    save_button = tk.Button(new_window, text="Save", command=save_use_case)
    save_button.pack()


def combine():
    # Function to handle the "Combine" button click event
    month = month_entry.get()
    year = year_entry.get()
    
    checkbox_functions = {
        "1": combine_CEMS(month, year),
        "2": combine_CHHA(month, year),
        "3": combine_claim_status(month, year),
        "4": combine_clinical_appeals(month, year),
        "5": combine_clinical_denials(month, year),
        "6": combine_IV(month, year),
        "7": combine_lab(month, year),
        "8": combine_minireg(month, year),
        "9": combine_prereg(month, year)
    }

    selected_checkboxes = [str(i) for i, var in enumerate(
        checkbox_vars, start=1) if var.get() == 1]
    for checkbox in selected_checkboxes:
        print(checkbox)
        if checkbox in checkbox_functions:
            checkbox_functions[checkbox]()


def render_checkboxes():
    # Function to render checkboxes based on data in the use_cases.json file
    with open("./use_cases.json", "r") as file:
        data = json.load(file)
        frame = tk.Frame(main_frame)
        frame.grid(row=2, column=0, columnspan=4)

        for i, use_case in enumerate(data, start=1):
            var = tk.IntVar()
            checkbox_vars.append(var)
            checkbox = tk.Checkbutton(
                frame, text=use_case["name"], variable=var)
            checkbox.grid(row=(i - 1) % 5, column=(i - 1) //
                          5, sticky="w", padx=10, pady=5)

# check if month and year are filled out
def check_entries(event):
    if month_entry.get() and year_entry.get():
        combine_button.config(state=tk.NORMAL)
    else:
        combine_button.config(state=tk.DISABLED)

root = tk.Tk()
root.geometry("400x300")
root.title("Consolidation Tool")

root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# frames
main_frame = tk.Frame(root)
main_frame.pack()

# Labels and Entries
tk.Label(main_frame, text="Enter month in MM format.").grid(
    row=0, column=0, columnspan=2)
month_entry = tk.Entry(main_frame)
month_entry.grid(row=0, column=2)
month_entry.bind("<KeyRelease>", check_entries)

tk.Label(main_frame, text="Enter year in YYYY format.").grid(
    row=1, column=0, columnspan=2)
year_entry = tk.Entry(main_frame)
year_entry.grid(row=1, column=2)
year_entry.bind("<KeyRelease>", check_entries)

# Checkboxes
checkbox_vars = []
render_checkboxes()

# "Add New Use Case" Button
add_use_case_button = tk.Button(
    main_frame, text="Add New Use Case", command=add_new_use_case)
add_use_case_button.grid(row=len(checkbox_vars) + 2, column=0)

# "Combine" Button
combine_button = tk.Button(main_frame, text="Combine", command=combine)
combine_button.grid(row=len(checkbox_vars) + 2, column=2, padx=10, pady=10)
combine_button.config(state=tk.DISABLED)



root.mainloop()
