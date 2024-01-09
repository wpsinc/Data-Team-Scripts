import pandas as pd
import os
from tqdm import tqdm
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
    filename for filename in os.listdir(StockStatus) if filename.endswith(".xlsx")
]

pbar = tqdm(total=len(xlsx_files))

# Loop through all files in the folder
for filename in xlsx_files:
    # Read the file into a dataframe
    file_df = pd.read_excel(os.path.join(StockStatus, filename))
    # Append the file dataframe to the main dataframe
    StockStatusDF = StockStatusDF._append(file_df)
    pbar.update(1)

pbar.close()

# Drop duplicates from the main dataframe
StockStatusDF.drop_duplicates(inplace=True)
StockStatusDF = StockStatusDF.rename(columns={"ItemNumber": "WPS Part Number"})
MegaDF = pd.read_excel(os.path.join(Mega, "Mega Report.xlsx"))
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
merged_df.columns
merged_df = merged_df.drop_duplicates()
merged_df.to_excel(
    os.path.join(Mega, "Mega Report.xlsx"),
    index=False,
)
