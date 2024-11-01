import pandas as pd
import os
import time
import chardet
from openpyxl import load_workbook
from halo import Halo
import warnings
from concurrent.futures import ThreadPoolExecutor

warnings.simplefilter("ignore")

# Get the current user's home directory
home_dir = os.path.expanduser("~")

# Replace hardcoded paths
base_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports"
)
eom_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates"
)
# Replace hardcoded paths
output_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Mega Report Output Reference"
)
def check_path(path, name):
    if not os.path.exists(path):
        print(f"Error: The path {path} does not exist for this user.")
        exit()
        
def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding']

check_path(base_path, "base_path")
check_path(eom_path, "eom_path")
check_path(output_path, "output_path")

# Navigate to folder containing Stock Status
StockStatus = os.path.join(base_path, "Stock Status")
check_path(StockStatus, "StockStatus")

# Navigate to folder containing inventory files
Mega = os.path.join(base_path, "Mega Report")

# Check if Mega exists
if not os.path.exists(Mega):
    print(f"Error: The path {Mega} does not exist for this user.")
    exit()

# Create an empty dataframe
StockStatusDF = pd.DataFrame()

# Get a list of all .csv files in the directory
csv_files = [entry.name for entry in os.scandir(StockStatus) if entry.is_file() and entry.name.endswith(".csv")]

dfs = []
spinner = Halo(text="Reading files...", spinner="dots")
spinner.start()

start_time = time.time()
start_time_operation = time.time()

def read_file(filename):
    file_path = os.path.join(StockStatus, filename)
    encoding = detect_encoding(file_path)
    # Read the file into a dataframe with the detected encoding and specified separator
    file_df = pd.read_csv(file_path, encoding=encoding, sep=",")
    return file_df

# Use ThreadPoolExecutor to read files in parallel
with ThreadPoolExecutor(max_workers=5) as executor:
    dfs = list(executor.map(read_file, csv_files))

# Concatenate all dataframes in the list
StockStatusDF = pd.concat(dfs, ignore_index=True)

# Use in-place operations
StockStatusDF.rename(
    columns={
        "ItemLookup_ItemNumber": "WPS Part Number",
        "Item_Detail_ItemStatus": "ItemStatus",
        "ItemLookup_Product_Manager": "Product Manager",
        "Item_Detail_VendorPartNumber": "VendorPartNumber",
        "ItemLookup_Description1": "Description1",
        "ItemLookup_Description2": "Description2",
        "Division_Division": "Division",
        "Class_Class": "Class",
        "SubClass_Sub_Class": "Sub-Class",
        "SubSubClass_Sub_Sub_Class": "Sub-Sub-Class",
        "Item_Flags_ModelDesc": "ModelDesc",
        "Item_Flags_StyleDesc": "StyleDesc",
        "Item_Flags_ColorDesc": "ColorDesc",
        "Item_Flags_SizeDescription": "SizeDescription",
        "Item_Flags_ApparelDesc": "ApparelDesc",
        "Item_Flags_YearDesign": "YearDesign",
        "Segment_Segment": "Segment",
        "SubSegment_Sub_Segment": "Sub-Segment",
        "Item_Detail_PreferredVendor": "PreferredVendor",
        "VendorDetail_Vendor": "Vendor",
        "Item_Detail_ItemCategory": "ItemCategory",
        "Brand_Lookup_Brand": "Brand",
        "ID_CalcTable_Prev_12_Months_QTY_Sold": "Total Prev 12 Units Sold",
        "ID_CalcTable_Prev_12_Months_COGS" : "Total Prev 12 COGS",
        "ID_CalcTable_Prev_12_Months_Sales": "Total Prev 12 Sales $",
        "ItemRank_ItemRank": "Item Rank",
        "Item_Detail_Cost": "Cost",
        "Item_Detail_DealerA": "Dealer A",
        "Item_Detail_DealerB": "Dealer B",
        "Item_Detail_ListPrice": "List Price"


    },
    
    inplace=True,
)

# Sort the DataFrame before dropping duplicates
StockStatusDF.sort_values("WPS Part Number", inplace=True)
StockStatusDF.drop_duplicates(inplace=True)
with open(os.path.join(Mega, "Mega Report.csv"), 'rb') as f:
    result = chardet.detect(f.read())
    encoding = result['encoding']
MegaDF = pd.read_csv(os.path.join(Mega, "Mega Report.csv"), encoding=encoding, sep="\t", header=0)

end_time_operation = time.time()
operation_duration = round((end_time_operation - start_time_operation) / 60, 5)
spinner.stop_and_persist(
    symbol="✔️ ".encode("utf-8"),
    text=f"Files Read. Process Time: {operation_duration} minutes",
)

MegaDF.sort_values("WPS Part Number", inplace=True)

spinner = Halo(text="Merging dataframes...", spinner="dots")
spinner.start()
start_time_operation = time.time()

merged_df = pd.merge(
    StockStatusDF[
        [
            "WPS Part Number",
            "ItemStatus",
            "Item Rank",
            "Product Manager",
            "VendorPartNumber",
            "Description1",
            "Description2",
            "Division",
            "Class",
            "Sub-Class",
            "Sub-Sub-Class",
            "ModelDesc",
            "StyleDesc",
            "ColorDesc",
            "SizeDescription",
            "ApparelDesc",
            "YearDesign",
            "Segment",
            "Sub-Segment",
            "PreferredVendor",
            "Vendor",
            "ItemCategory",
            "Brand",
            "Cost",
            "Dealer A",
            "Dealer B",
            "List Price",
            "Total Prev 12 Units Sold",
            "Total Prev 12 COGS",
            "Total Prev 12 Sales $"
        ]
    ],
    MegaDF,
    on="WPS Part Number",
    how="right",
)
# Sort and drop duplicates after merging
merged_df.sort_values("WPS Part Number", inplace=True)
merged_df.drop_duplicates(inplace=True)

end_time_operation = time.time()
operation_duration = round((end_time_operation - start_time_operation) / 60, 5)
spinner.stop_and_persist(
    symbol="✔️ ".encode("utf-8"),
    text=f"Dataframes Merged. Process Time: {operation_duration} minutes",
)

spinner = Halo(text="Cleaning file...", spinner="dots")
spinner.start()
start_time_operation = time.time()

merged_df = merged_df[merged_df["WPS Part Number"].notnull()]

merged_df.to_excel(os.path.join(output_path, "Mega Report Output.xlsx"), index=False)
end_time_operation = time.time()
operation_duration = round((end_time_operation - start_time_operation) / 60, 5)
spinner.stop_and_persist(
    symbol="✔️ ".encode("utf-8"),
    text=f"File Cleaned. Process Time: {operation_duration} minutes",
)

end_time = time.time()
print(f"Merging took {round((end_time - start_time)/60, 5)} minutes")