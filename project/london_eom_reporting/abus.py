import pandas as pd
import uuid
import os
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
s2k_sales = filedialog.askopenfilename(title="Select S2K Analytic Sales File", filetypes=[("Excel Files", "*.xlsx;*.xls")])
stock_status = filedialog.askopenfilename(title="Select Stock Status File", filetypes=[("Excel Files", "*.xlsx;*.xls")])

try:
    data1 = pd.read_excel(s2k_sales, dtype=str)
    data2 = pd.read_excel(stock_status, dtype=str)
except:
    data1 = pd.read_csv(s2k_sales, dtype=str, encoding='latin1')
    data2 = pd.read_csv(stock_status, dtype=str, encoding='latin1')


data1 = data1.drop_duplicates()
data2.rename(columns={'Item No.': 'WPS Part Number'}, inplace=True)
output = pd.merge(data1, data2,
                   on='WPS Part Number',
                   how='left')


output = output.loc[:,~output.columns.duplicated()]

# Create output directory if it doesn't exist
output_path = os.path.join(os.getcwd(), 'output')
if not os.path.exists(output_path):
    os.makedirs(output_path)

# Generate a unique filename using uuid
filename = f"output_{str(uuid.uuid4().hex[:8])}.csv"

# Write output to csv file with relative path
output_file = os.path.join(output_path, filename)
output.to_csv(output_file, index=False)

print(f"Output written to {filename}")


print(output.shape[0])
print(data1.shape[0])