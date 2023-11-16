import pandas as pd
import uuid
import os
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
sales_file_path = filedialog.askopenfilename(title="Select Sales file", filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx;*.xls")])
item_details_file_path = filedialog.askopenfilename(title="Select Item Details file", filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx;*.xls")])
cross_drop_file_path = filedialog.askopenfilename(title="Select Cross and Drop file", filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx;*.xls")])

data1 = pd.read_csv(sales_file_path, dtype=str)
data2 = pd.read_csv(item_details_file_path, dtype=str)
data3 = pd.read_csv(cross_drop_file_path, dtype=str)

data1 = data1.drop_duplicates()
data1.rename(columns={'ItemNum': 'ItemNumber'}, inplace=True)
data3.rename(columns={'Invoice': 'InvoiceNum'}, inplace=True)
Preoutput = pd.merge(data1, data2,
                   on='ItemNumber',
                   how='left')

output = pd.merge(Preoutput, data3,
                  on='InvoiceNum',
                   how='left')

# Change Ord & Warehouse output
output = output.rename(columns={
    'Ord-Line': 'OrdLine',
    'Full Order #': 'Order #',
    'Ord': 'Item Number',
    'InvoiceDate': 'Invoice Date',
    'LineSales': 'Line Sales $',
    'COGS': 'COGS $',
    'CustomerNumber': 'Customer Number',
    'Warehouse': 'Ship Warehouse',
    'InvoiceNum': 'InvoiceNumber',
    'Description1': 'Description 1',
    'Description2': 'Description 2',
    'Cross Ship Flag': 'CrossShip Flag',
    'Drop Ship Flag': 'DropShip Flag',
})

output = output[[
    'OrdLine',
    'Order #',
    'Line',
    'ItemNumber',
    'Line Sales $',
    'COGS $',
    'ShipQty',
    'Invoice Date',
    'Customer Number',
    'Ship Warehouse',
    'InvoiceNumber',
    'CrossShip Flag',
    'DropShip Flag',
    'Description 1',
    'Description 2',
    'Brand',
    'Segment',
    'Division',
    'Class',
    'Sub-Class',
    'Sub-Sub-Class'
    ]]
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