import pandas as pd
import os

current_dir = os.path.dirname(os.path.abspath(__file__))

input_folder = os.path.join(current_dir, 'WPS_Harley')
output_file = os.path.join(input_folder, 'output.xlsx')

#input_folder = 'C:\Users\LucyHaskew\Downloads\WPS Harley'
#output_file = 'C:\Users\LucyHaskew\Downloads\WPS Harley\output.xlsx'
common_start = 'WPS HARLEY DAVIDSON'

dataframes = []

for files in os.listdir(input_folder):
    if files.startswith(common_start):
        file_path = os.path.join(input_folder, files)
        read_files = pd.read_excel(file_path)

        dataframes.append()

combined_read_files = pd.concat(dataframes, ignore_index=True, sort=False)

combined_read_files.to_excel(output_file, index=False)