import pandas as pd
import os

input_folder = 'C:/Users/LucyHaskew/Downloads/WPS_Harley'
output_file = 'C:/Users/LucyHaskew/Downloads/WPS_Harley/output.xlsx'
common_end = '.xlsx'
 
dataframes = []

for files in os.listdir(input_folder):
    if files.endswith(common_end):
        file_path = os.path.join(input_folder, files)
        read_files = pd.read_excel(file_path)
        read_files = read_files.dropna(axis=1, how='all')

        dataframes.append(read_files)

combined_read_files = pd.concat(dataframes, axis=1)

combined_read_files.to_excel(output_file, index=False)