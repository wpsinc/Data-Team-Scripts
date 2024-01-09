import pandas as pd
import os
from tqdm import tqdm


# Navigate to folder containing Stock Status
StockStatus = "C:/Users/megan.partridge/OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Stock Status"
# Navigate to folder containing inventory files
Mega = "C:/Users/megan.partridge/OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Mega Report"

# Create an empty dataframe
StockStatusDF = pd.DataFrame()

# Loop through all files in the folder
for filename in os.listdir(StockStatus):
    if filename.endswith('.xlsx'):
        # Read the file into a dataframe
        file_df = pd.read_excel(os.path.join(StockStatus, filename))
        # Append the file dataframe to the main dataframe
        StockStatusDF = StockStatusDF._append(file_df)
# Drop duplicates from the main dataframe
StockStatusDF.drop_duplicates(inplace=True)
StockStatusDF = StockStatusDF.rename(columns={'ItemNumber': 'WPS Part Number'})
MegaDF = pd.read_excel(os.path.join(Mega, 'Mega Report.xlsx'))
merged_df = pd.merge(StockStatusDF[['WPS Part Number', 'ItemStatus', 'Product Manager', 'OEMPartNumber', 'Description1', 'Description2', 'Division', 'Class', 'Sub-Class', 'Sub-Sub-Class', 'ModelDesc', 'StyleDesc', 'ColorDesc', 'SizeDescription', 'ApparelDesc', 'YearDesign', 'Segment', 'Sub-Segment', 'PreferredVendor', 'Vendor', 'ItemCategory', 'Brand']], MegaDF, on='WPS Part Number', how='inner')
merged_df.columns
merged_df = merged_df.drop_duplicates()
merged_df.to_excel('C:/Users/megan.partridge/OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Mega Report.xlsx', index = False)
with tqdm(total=merged_df.shape[0]) as pbar:
    for index, row in merged_df.iterrows():
        merged_df.at[index, 'A'] = row['A'] * 2
        pbar.update(1)
