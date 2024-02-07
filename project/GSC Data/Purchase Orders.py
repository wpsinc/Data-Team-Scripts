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
ReceiptsDF = pd.read_csv(
    Receipts_file, dtype=str, encoding=Receipts_encoding, sep="\t", engine="python"
)
ReceiptsDF.rename(
    columns={
        "P.O. #": "Purchase Order Number",
        "P.O. Line": "Purchase Order Line Number",
        "Item Code": "Part Number",
    },
    inplace=True,
)
spinner.stop_and_persist(symbol="✔️ ".encode("utf-8"), text="Files Read")

spinner = Halo(text="Cleaning file...", spinner="dots")

spinner.start()
start_time = time.time()

Purchases_file = os.path.join(Purchases, "Purchase Order Data Karen.csv")
Purchases_encoding = find_encoding(Purchases_file)
PurchasesDF = pd.read_csv(
    Purchases_file, dtype=str, encoding=Purchases_encoding, sep="\t", engine="python"
)

print("PurchasesDF: ", PurchasesDF.columns)
print(ReceiptsDF.columns)
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
merged_df.rename(
    columns={"Order Date_x": "Order Date", "Part Number_x": "Part Number"}, inplace=True
)
merged_df.drop(["Part Number_y", "Order Date_y"], axis=1, inplace=True)
merged_df.drop_duplicates(inplace=True)
# convert the columns to numeric
merged_df["Quantity Ordered"] = pd.to_numeric(
    merged_df["Quantity Ordered"], errors="coerce"
)
merged_df["Unit Price"] = pd.to_numeric(merged_df["Unit Price"], errors="coerce")
merged_df["Total Spend USD"] = pd.to_numeric(
    merged_df["Total Spend USD"], errors="coerce"
)
# Fill Na to 0
merged_df.fillna(0, inplace=True)
# Create Net Total
merged_df["Net Total"] = merged_df["Quantity Ordered"] * merged_df["Unit Price"]
# Sum Net Total
merged_df["Net Total Sum"] = merged_df.groupby("Purchase Order Number")[
    "Net Total"
].transform("sum")
# calculate the discount percentage
merged_df["Discount Percentage"] = (
    (merged_df["Total Spend USD"] - merged_df["Net Total Sum"])
    / merged_df["Total Spend USD"]
) * 100
merged_df["New Total Spend USD"] = merged_df["Net Total"] * (
    1 - (merged_df["Discount Percentage"] / 100)
)

merged_df.drop(
    [
        "Discount Percentage",
        "Total Spend USD",
        "Total Spend Local",
        "Net Total Sum",
        "Net Total",
    ],
    axis=1,
    inplace=True,
)
merged_df["Total Spend Local"] = merged_df["New Total Spend USD"]
merged_df.rename(columns={"New Total Spend USD": "Total Spend USD"}, inplace=True)

new_order = [
    "Company ID",
    "Purchase Order Number",
    "Purchase Order Line Number",
    "Purchase Order Date",
    "Order Date",
    "Supplier Master ID",
    "Supplier Name",
    "Part Number",
    "Part Name",
    "Part Description",
    "Unit of Measure",
    "Quantity Ordered",
    "Quantity Received",
    "Unit Price",
    "Total Spend Local",
    "Total Spend USD",
    "Currency Code",
    "Facility ID",
    "Facility",
    "Received Date",
    "Buyer Number",
    "Buyer Name",
    "Payment Terms Code",
    "Payment Terms definition",
    "Facility Address 1",
    "Facility City",
    "Facility State/Province",
    "Facility Country",
]
# convert zeros to None in the 'Received Date' column
merged_df.loc[merged_df['Received Date'] == '0', 'Received Date'] = None

# convert the date columns to datetime format
merged_df["Order Date"] = pd.to_datetime(merged_df["Order Date"], errors='coerce')
merged_df["Purchase Order Date"] = pd.to_datetime(merged_df["Purchase Order Date"], errors='coerce')
merged_df["Received Date"] = pd.to_datetime(merged_df["Received Date"], errors='coerce')

# reformat the date columns
merged_df["Order Date"] = merged_df["Order Date"].dt.strftime("%m/%d/%Y", errors='coerce')
merged_df["Purchase Order Date"] = merged_df["Purchase Order Date"].dt.strftime("%m/%d/%Y", errors='coerce')
merged_df["Received Date"] = merged_df["Received Date"].dt.strftime("%m/%d/%Y", errors='coerce')


merged_df = merged_df.reindex(columns=new_order)
merged_df.to_csv(
    os.path.join(output_path, "Purchase Orders.csv"),
    index=False,
)

spinner.stop_and_persist(symbol="✔️ ".encode("utf-8"), text=" File Cleaned")

end_time = time.time()
print(f"Merging took {round((end_time - start_time)/60, 5)} minutes.")