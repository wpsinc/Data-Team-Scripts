import pandas as pd
import tkinter as tk
import os
from tkinter import filedialog

# Get the current user's home directory
home_dir = os.path.expanduser("~")

# Define the path for the output file
demand_file_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/Demand.xlsx"
)

# Function to get the file path using a file dialog
def get_input_file_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(defaultextension=".xlsm", filetypes=[("Excel files", "*.xlsm")])
    return file_path

# Read data from the Purchasing tab
with pd.ExcelFile(get_input_file_path()) as xlsm:
    df_existing = pd.read_excel(xlsm, sheet_name="Purchasing")

# Read data from the Demand file
with pd.ExcelFile(demand_file_path) as xlsx:
    Demand = pd.read_excel(xlsx, sheet_name="Purchasing")

# Append new records to existing data
Demand = pd.concat([Demand, df_existing], ignore_index=True)

# Save the combined data to the specified output file
Demand.to_excel(demand_file_path, sheet_name="Purchasing", index=False)

print(f"Records appended successfully to {demand_file_path}!")