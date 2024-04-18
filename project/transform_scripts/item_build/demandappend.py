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
df_existing = pd.read_excel(get_input_file_path(), sheet_name="Purchasing")

# Create new records (example data) from the Purchasing tab
# You can modify this part to include actual data
new_records = df_existing.copy()  # Example: Copy existing data

# Append new records to existing data
df_combined = pd.concat([df_existing, new_records], ignore_index=True)

# Save the combined data to the specified output file
df_combined.to_excel(demand_file_path, sheet_name="Purchasing", index=False)

print(f"Records appended successfully to {demand_file_path}!")
