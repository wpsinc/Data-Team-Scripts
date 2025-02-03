import pandas as pd
import os
import shutil
from datetime import datetime
import chardet

# Get the current user's home directory
home_dir = os.path.expanduser("~")
reports_dir = os.path.join(home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates/SpecialVendorReports")

vendor_dirs = os.listdir(reports_dir)

for vendor_dir in vendor_dirs:
    src_dir = os.path.join(reports_dir, vendor_dir)
    dst_dir = os.path.join(src_dir, "Report History")

    files = [f for f in os.listdir(src_dir) if os.path.isfile(os.path.join(src_dir, f))]

    if not files:
        dir_name = os.path.basename(src_dir)
        print(f"No files to process for {dir_name}")
        continue

    inventory_file = None
    sales_file = None

    for file in files:
        if file.endswith('Inventory.csv'):
            inventory_file = os.path.join(src_dir, file)
        elif file.endswith('Sales.csv'):
            sales_file = os.path.join(src_dir, file)

    if inventory_file and sales_file:
        with open(inventory_file, 'rb') as f:
            result = chardet.detect(f.read())
            encoding = result['encoding']

        inventory_df = pd.read_csv(inventory_file, sep="\t", encoding=encoding)
        sales_df = pd.read_csv(sales_file, usecols=['ItemLookup_ItemNumber', 'DateDimension_Month_Started', 'ID_CalcTable_Sales_Qty', 'ID_CalcTable_Sales_'])

        # Rename columns for consistency
        sales_df.rename(columns={'ItemLookup_ItemNumber': 'ItemNumber', 'DateDimension_Month_Started': 'SalesDate', 'ID_CalcTable_Sales_Qty': 'SalesQty', 'ID_CalcTable_Sales_': 'Sales $'}, inplace=True)

        # Convert Sales Date to datetime
        sales_df['SalesDate'] = pd.to_datetime(sales_df['SalesDate'])

        # Create a date range for pivoting
        date_range = pd.date_range(start='2024-01-01', end='2025-12-01', freq='MS')

        # Pivot the table
        pivot_df = sales_df.pivot_table(index='ItemNumber', columns='SalesDate', values=['SalesQty', 'Sales $'], aggfunc='sum', fill_value=0)

        # Ensure all dates in the range are included
        pivot_df = pivot_df.reindex(columns=pd.MultiIndex.from_product([['SalesQty', 'Sales $'], date_range]), fill_value=0)

        # Flatten the MultiIndex columns
        pivot_df.columns = ['_'.join([str(col) for col in col_tuple]) for col_tuple in pivot_df.columns]

        # Merge the dataframes on ItemNumber
        merged_df = pd.merge(inventory_df, pivot_df, left_on='WPS Part Number', right_on='ItemNumber', how='left')

        # Remove the Inventory and Sales files
        os.remove(inventory_file)
        os.remove(sales_file)

        # Save the pivot table to a new CSV file
        output_file = os.path.join(src_dir, f"{vendor_dir} Sales.csv")
        merged_df.to_csv(output_file, index=False)

        try:
            year, month, _ = file.split("_")[1].split("-")
        except IndexError:
            creation_date = datetime.fromtimestamp(os.path.getctime(os.path.join(src_dir, file))).strftime("%Y-%m-%d")
            new_file = f"{vendor_dir}_{creation_date}.{file.split('.')[1]}"
            os.rename(os.path.join(src_dir, file), os.path.join(src_dir, new_file))
            file = new_file
            year, month, _ = creation_date.split("-")

        year_dir = os.path.join(dst_dir, year)

        if not os.path.exists(year_dir):
            os.makedirs(year_dir)
            print(f"Created new directory: {year_dir}")

        month_dir = os.path.join(year_dir, f"{month}-{year[2:]}")

        if not os.path.exists(month_dir):
            os.makedirs(month_dir)
            print(f"Created new directory: {month_dir}")

        try:
            shutil.move(os.path.join(src_dir, file), os.path.join(month_dir, file))
            print(f"Successfully moved {file} to {vendor_dir}_{month}-{year[2:]}")
        except Exception as e:
            print(f"Error moving {file} to {month_dir}: {e}")
