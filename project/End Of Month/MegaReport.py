import pandas as pd
import os
import time
from halo import Halo
import warnings

warnings.simplefilter("ignore")

# Get the current user's home directory
home_dir = os.path.expanduser("~")

# Replace hardcoded paths
base_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports"
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
    entry.name for entry in os.scandir(StockStatus) if entry.is_file() and entry.name.endswith(".xlsx")
]

dfs = []

spinner = Halo(text="Reading files...", spinner="dots")
spinner.start()

start_time = time.time()

# Loop through all files in the folder
for filename in xlsx_files:
    # Read the file into a dataframe
    file_df = pd.read_excel(os.path.join(StockStatus, filename), engine="openpyxl")
    # Append the dataframe to the list
    dfs.append(file_df)

# Concatenate all dataframes in the list
StockStatusDF = pd.concat(dfs)


# Rename the column 'ItemNumber' to 'WPS Part Number'
StockStatusDF = StockStatusDF.rename(columns={"ItemNumber": "WPS Part Number"})

# Convert 'WPS Part Number' column in StockStatusDF to string type
StockStatusDF["WPS Part Number"] = StockStatusDF["WPS Part Number"].astype(str)

# Drop duplicates from the main dataframe
StockStatusDF.drop_duplicates(inplace=True)
MegaDF = pd.read_excel(os.path.join(Mega, "Mega Report.xlsx"), sheet_name="page")

spinner.stop_and_persist(symbol="✔️".encode("utf-8"), text=" Files Read")

spinner = Halo(text="Merging dataframes...", spinner="dots")
spinner.start()

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
    how="inner",
)

spinner.stop_and_persist(symbol="✔️".encode("utf-8"), text=" Dataframes Merged")

spinner = Halo(text="Cleaning file...", spinner="dots")
spinner.start()
merged_df = merged_df.drop_duplicates(inplace=True)
merged_df.to_excel(
    r"C:\Users\London.Perry\OneDrive - Arrowhead EP\Data Tech\End of Month Templates\Linked Reports\Mega Report\Mega Report1.xlsx",
    index=False,
)
spinner.stop_and_persist(symbol="✔️".encode("utf-8"), text=" File Cleaned")

end_time = time.time()
print(f"Merging took {round((end_time - start_time)/60, 5)} minutes.")
