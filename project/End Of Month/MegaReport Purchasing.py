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
    home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Mega Report Purchasing Output Reference"
)
def check_path(path, name):
    if not os.path.exists(path):
        print(f"Error: The path {path} does not exist for this user.")
        exit()

# Check paths
check_path(base_path, "base_path")
check_path(eom_path, "eom_path")
check_path(output_path, "output_path")
# Navigate to folder containing Stock Status
StockStatus = os.path.join(base_path, "Stock Status")
check_path(StockStatus, "StockStatus")

# Navigate to folder containing inventory files
Mega = os.path.join(base_path, "Mega Report Purchasing")

# Check if Mega exists
if not os.path.exists(Mega):
    print(f"Error: The path {Mega} does not exist for this user.")
    exit()

# Create an empty dataframe
StockStatusDF = pd.DataFrame()

# Get a list of all .xlsx files in the directory
xlsx_files = [entry.name for entry in os.scandir(StockStatus) if entry.is_file() and entry.name.endswith(".xlsx")]

dfs = []
spinner = Halo(text="Reading files...", spinner="dots")
spinner.start()

start_time = time.time()
start_time_operation = time.time()

def read_file(filename):
    # Read the file into a dataframe
    file_df = pd.read_excel(os.path.join(StockStatus, filename), engine="openpyxl")
    return file_df

# Use ThreadPoolExecutor to read files in parallel
with ThreadPoolExecutor(max_workers=5) as executor:
    dfs = list(executor.map(read_file, xlsx_files))

# Concatenate all dataframes in the list
StockStatusDF = pd.concat(dfs, ignore_index=True)


# Use in-place operations
StockStatusDF.rename(
    columns={
        "ItemLookup[ItemNumber]": "WPS Part Number",
        "Item Detail[ItemStatus]": "ItemStatus",
        "ItemLookup[Product Manager]": "Product Manager",
        "Item Detail[OEMPartNumber]": "OEMPartNumber",
        "ItemLookup[Description1]": "Description1",
        "ItemLookup[Description2]": "Description2",
        "Division[Division]": "Division",
        "Class[Class]": "Class",
        "SubClass[Sub-Class]": "Sub-Class",
        "SubSubClass[Sub-Sub-Class]": "Sub-Sub-Class",
        "Item Flags[ModelDesc]": "ModelDesc",
        "Item Flags[StyleDesc]": "StyleDesc",
        "Item Flags[ColorDesc]": "ColorDesc",
        "Item Flags[SizeDescription]": "SizeDescription",
        "Item Flags[ApparelDesc]": "ApparelDesc",
        "[SumYearDesign]": "YearDesign",
        "Segment[Segment]": "Segment",
        "SubSegment[Sub-Segment]": "Sub-Segment",
        "Item Detail[PreferredVendor]": "PreferredVendor",
        "VendorDetail[Vendor]": "Vendor",
        "Item Detail[ItemCategory]": "ItemCategory",
        "Brand Lookup[Brand]": "Brand",
    },
    inplace=True,
)
# Sort the DataFrame before dropping duplicates
StockStatusDF.sort_values("WPS Part Number", inplace=True)
StockStatusDF.drop_duplicates(inplace=True)
with open(os.path.join(Mega, "Mega Report Purchasing.csv"), 'rb') as f:
    result = chardet.detect(f.read())
    encoding = result['encoding']
MegaDF = pd.read_csv(os.path.join(Mega, "Mega Report Purchasing.csv"), encoding=encoding, sep="\t", header=0)

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
            "Product Manager",
            "OEMPartNumber",
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

merged_df.to_excel(os.path.join(output_path, "Mega Report Purchasing Output.xlsx"), index=False)
end_time_operation = time.time()
operation_duration = round((end_time_operation - start_time_operation) / 60, 5)
spinner.stop_and_persist(
    symbol="✔️ ".encode("utf-8"),
    text=f"File Cleaned. Process Time: {operation_duration} minutes",
)

end_time = time.time()
print(f"Merging took {round((end_time - start_time)/60, 5)} minutes")