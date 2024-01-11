import pandas as pd
import os
import time
from halo import Halo
import warnings
import chardet

warnings.simplefilter("ignore")

def find_encoding(fname):
    r_file = open(fname, "rb").read()
    result = chardet.detect(r_file)
    charenc = result["encoding"]
    return charenc

# Get the current user's home directory
home_dir = os.path.expanduser("~")

# Replace hardcoded paths
base_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Teamwork Data/2024-01/Karen Project PO Data"
)
output_path = os.path.join(
    home_dir,
    "OneDrive - Arrowhead EP/Teamwork Data/2024-01/Karen Project PO Data/Purchase Order Data",
)

# Check if base_path exists
if not os.path.exists(base_path):
    print(f"Error: The path {base_path} does not exist for this user.")
    exit()

# Navigate to folder containing inventory files
Receipts = os.path.join(base_path, "Receipts")

# Check if Mega exists
if not os.path.exists(Receipts):
    print(f"Error: The path {Receipts} does not exist for this user.")
    exit()

# Navigate to folder containing inventory files
Purchases = os.path.join(base_path, "Purchases")

# Check if Mega exists
if not os.path.exists(Purchases):
    print(f"Error: The path {Purchases} does not exist for this user.")
    exit()


spinner = Halo(text="Reading files...", spinner="dots")
spinner.start()

start_time = time.time()

Receipts_file = os.path.join(Receipts, "P.O. Receipts Karen.csv")
Receipts_encoding = find_encoding(Receipts_file)
ReceiptsDF = pd.read_csv(Receipts_file, dtype=str, encoding=Receipts_encoding, engine='python')

spinner.stop_and_persist(symbol="✔️ ".encode("utf-8"), text="Files Read")

spinner = Halo(text="Cleaning file...", spinner="dots")

spinner.start()
start_time = time.time()

Purchases_file = os.path.join(Purchases, "Purchase Order Data Karen.csv")
Purchases_encoding = find_encoding(Purchases_file)
PurchasesDF = pd.read_csv(Purchases_file, dtype=str, encoding=Purchases_encoding, engine='python')
PurchasesDF.rename(
    columns={
        "P.O. #": "Purchase Order Number",
        "P.O. Line": "Purchase Order Line Number",
        "Item Code": "Part Number",
    },
    inplace=True,
)

spinner.stop_and_persist(symbol="✔️ ".encode("utf-8"), text="Files Read")


spinner = Halo(text="Merging dataframes...", spinner="dots")
spinner.start()

merged_df = pd.merge(
    PurchasesDF,
    ReceiptsDF,
    how="left",
    on=["Purchase Order Number", "Purchase Order Line Number"],
)
spinner.stop_and_persist(symbol="✔️ ".encode("utf-8"), text=" Dataframes Merged")

spinner = Halo(text="Cleaning file...", spinner="dots")
spinner.start()

merged_df.drop_duplicates(inplace=True)
merged_df.to_csv(
    os.path.join(output_path, "Purchase Orders.csv"),
    index=False,
)

spinner.stop_and_persist(symbol="✔️ ".encode("utf-8"), text=" File Cleaned")

end_time = time.time()
print(f"Merging took {round((end_time - start_time)/60, 5)} minutes.")