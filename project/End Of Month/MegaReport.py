import pandas as pd
import os
import time
from halo import Halo
from concurrent.futures import ThreadPoolExecutor
import win32com.client

# Get the current user's home directory
home_dir = os.path.expanduser("~")

# Replace hardcoded paths
base_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports"
)
eom_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates"
)
output_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Mega Report Output Reference"
)

# Check if base_path exists
if not os.path.exists(base_path):
    print(f"Error: The path {base_path} does not exist for this user.")
    exit()

# Navigate to folder containing Stock Status
StockStatus = os.path.join(base_path, "Stock Status")

# Check if StockStatus exists
if not os.path.exists(StockStatus):
    print(f"Error: The path {StockStatus} does not exist for this user.")
    exit()

# Navigate to folder containing inventory files
Mega = os.path.join(base_path, "Mega Report")

# Check if Mega exists
if not os.path.exists(Mega):
    print(f"Error: The path {Mega} does not exist for this user.")
    exit()

# Create an empty dataframe
StockStatusDF = pd.DataFrame()

# Get a list of all .xlsx files in the directory
xlsx_files = [
    entry.name
    for entry in os.scandir(StockStatus)
    if entry.is_file() and entry.name.endswith(".xlsx")
]

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
StockStatusDF = pd.concat(dfs)

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

MegaDF = pd.read_excel(os.path.join(Mega, "Mega Report.xlsx"), sheet_name="page")

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

merged_df.to_excel(os.path.join(output_path, "Mega Report Output.xlsx"), index=False)
end_time_operation = time.time()
operation_duration = round((end_time_operation - start_time_operation) / 60, 5)
spinner.stop_and_persist(
    symbol="✔️ ".encode("utf-8"),
    text=f"File Cleaned. Process Time: {operation_duration} minutes",
)

spinner = Halo(text="Refreshing File Queries...", spinner="dots")
spinner.start()
start_time_operation = time.time()

# Define the output file
output_file = os.path.join(eom_path, "Mega Report Output.xlsx")

# Open Excel
Excel = win32com.client.Dispatch("Excel.Application")
Excel.Visible = False  # Excel will run in the background

# Open workbook
wb = Excel.Workbooks.Open(output_file)

# Refresh all data connections.
wb.RefreshAll()

# Save and close
wb.Save()
wb.Close()

# Quit Excel
Excel.Quit()

end_time_operation = time.time()
operation_duration = round((end_time_operation - start_time_operation) / 60, 5)
spinner.stop_and_persist(
    symbol="✔️ ".encode("utf-8"),
    text=f"File Queries Refreshed. Process Time: {operation_duration} minutes",
)

end_time = time.time()
print(f"Merging took {round((end_time - start_time)/60, 5)} minutes")
