import os
import json
import pandas as pd

# Get the current user's home directory.
home_dir = os.path.expanduser("~")

# Designated paths.
file_path = os.path.join(
    home_dir, f"OneDrive - Arrowhead EP/Data Tech/How To Templates/API Templates/API Files/Vehicle Breakdown.json"
)
output_file_path = os.path.join(
    home_dir, f"OneDrive - Arrowhead EP/Data Tech/How To Templates/API Templates/VehicleBreakdown.csv"
)
vehicles_file = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/Fitment Audit/Vehicle File/Vehicles.xlsx"
)

# Reads the JSON file.
with open(file_path, "r") as json_file:
    data = json.load(json_file)

# Extracts vehicle data.
vehicle_data = data["data"]
rows = []

# Processes the JSON data.
for entry in vehicle_data:
    if "vehicles" in entry and "data" in entry["vehicles"] and entry["vehicles"]["data"]:
        product_id = entry["id"]
        sku = str(entry["sku"])  # Convert to string
        for v in entry["vehicles"]["data"]:
            vehicle_id = str(v["id"]).strip() if "id" in v else None
            if vehicle_id and vehicle_id != "-":
                rows.append((sku, product_id, vehicle_id))

# Creates a DataFrame.
VehicleBreakdown = pd.DataFrame(rows, columns=["SKU", "Product_ID", "Vehicle_ID"])

# Ensures SKU is treated as a string and wrapped in quotes.
VehicleBreakdown["SKU"] = '"' + VehicleBreakdown["SKU"].astype(str) + '"'

# Clean and format data.
VehicleBreakdown["Vehicle_ID"] = pd.to_numeric(VehicleBreakdown["Vehicle_ID"], errors="coerce")
VehicleBreakdown["Product_ID"] = VehicleBreakdown["Product_ID"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
VehicleBreakdown["URL"] = "https://acp.wps-inc.com/items/" + VehicleBreakdown["Product_ID"] + "/edit"

# Loads and merges vehicle details.
vehicles_df = pd.read_excel(vehicles_file)
merged_df = pd.merge(VehicleBreakdown, vehicles_df, left_on="Vehicle_ID", right_on="vehicle_id", how="inner")

# Select required columns.
selected_columns = ["URL", "SKU", "vehicle_type", "year", "make", "model"]
filtered_df = merged_df[selected_columns]

filtered_df.to_csv(output_file_path, index=False)
print(f"Output saved to {output_file_path}")
