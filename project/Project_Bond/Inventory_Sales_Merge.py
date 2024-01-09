import pandas as pd
import os
import shutil

# Navigate to folder containing prev 12 months files
prev12 = "C:/Users/megan.partridge/OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Prev 12 Months Sales"
# Navigate to folder containing inventory files
inv = "C:/Users/megan.partridge/OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Inventory Reports"
prev12_processed = "C:/Users/megan.partridge/OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Prev 12 Months Sales Processed"
inv_processed = "C:/Users/megan.partridge/OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Inventory Reports Processed" 
# destination folder
destination = "C:/Users/megan.partridge/OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Inventory Project"

def merg_files(inv, prev12, output_dir):
    # Ensure output directory exists
    if not os.path.exists(destination):
        os.makedirs(destination)

    # Get list of common files in both directories
    invfiles = set(os.listdir(inv))
    prev12files = set(os.listdir(prev12))
    common_files = invfiles.intersection(prev12files)

    print(common_files)  # print the common files

    # Merge common files
    for filename in common_files:
        try:
            # Read the files into data frames
            Inventory = pd.read_excel(os.path.join(inv, filename))
            Sales =  pd.read_excel(os.path.join(prev12, filename))

            # Rename 'Item Code' to 'ItemNum' in Inventory
            Inventory = Inventory.rename(columns={'Item Code': 'ItemNum'})
        
            print(Sales.columns)
            print(Inventory.columns)
            # Merge on ItemNum
            merged_df = pd.merge(Inventory, Sales, on='ItemNum', suffixes=('_inv', '_prev12'))


            # Aggregate 'Sellable QTY' at the index level
            merged_df['Total Sellable QTY'] = merged_df.groupby(['ItemNum', 'Prev 12 Months Sales QTY', 'ICITMSTS'])['Sellable QTY'].transform('sum')

            # Create the pivot table with 'OH $' as values
            pivoted_df = merged_df.pivot_table(values='OH $', 
                                   index=['ItemNum', 'Prev 12 Months Sales QTY', 'ICITMSTS', 'Total Sellable QTY'], 
                                   columns='Location', 
                                   aggfunc='sum')

            # Write the pivoted dataframe to a file in the output directory
            pivoted_df.to_excel(os.path.join(destination, filename), index=True)
        except Exception as e:
            print(f"Failed to merge file {filename}. Reason: {str(e)}")

# Usage
merg_files(inv, prev12, destination)

for file_name in os.listdir(prev12):
    shutil.move(os.path.join(prev12, file_name), prev12_processed)
for file_name in os.listdir(inv):
    shutil.move(os.path.join(inv, file_name), inv_processed)    