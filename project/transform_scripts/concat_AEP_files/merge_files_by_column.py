import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory

def main():
    Tk().withdraw()
    files_dir = askdirectory(title="Select the folder with AEP catalog files you'd like to combine")
    print(f"Folder Path: {files_dir}")
    return files_dir

input_folder = main()
output_folder = os.path.join(input_folder, "Output")

# Create 'Output' directory if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

output_file = os.path.join(output_folder, "output.xlsx")
common_end = '.xlsx'

dataframes = []

for file in os.listdir(input_folder):
    if file.endswith(common_end):
        file_path = os.path.join(input_folder, file)
        read_file = pd.read_excel(file_path, skiprows=2)
        dataframes.append(read_file)

combined_read_files = pd.concat(dataframes, ignore_index=True)

# Unpivot the columns to the right of 'Years'
common_vars = ['Type', 'Make', 'CC', 'Model', 'Years']
value_vars = combined_read_files.columns.difference(common_vars)

melted_df = combined_read_files.melt(id_vars=common_vars, value_vars=value_vars, var_name='Attribute', value_name='Value')

# Drop rows where 'Value' is NaN
melted_df = melted_df.dropna(subset=['Value'])

if not melted_df.empty:
    melted_df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")
else:
    print("Error: 'melted_df' is empty. Check your melting logic.")
