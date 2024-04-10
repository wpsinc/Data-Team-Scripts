import pandas as pd
import json

# Specify the file path
file_path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\Fitment Json\Online Fitment.json"
output_file_path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\onlinefitment.csv"
Fitment_path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Online Fitment.json"
# Read the JSON data
with open(file_path, "r") as json_file:
    data = json.load(json_file)

# Extract relevant information from the "vehicles" section (if available)
vehicle_data = data["data"]
rows = []

for entry in vehicle_data:
    if "vehicles" in entry and "data" in entry["vehicles"] and entry["vehicles"]["data"]:
        # Assuming you want all vehicle IDs next to the corresponding SKU
        vehicle_ids = [str(v["id"]) for v in entry["vehicles"]["data"] if v["id"] and v["id"].strip() != "-"]  # Exclude empty IDs and hyphens
        sku = str(entry["sku"])  # Convert to string
        for vehicle_id in vehicle_ids:
            rows.append((sku, vehicle_id))
    else:
        # Handle cases where no vehicle data is available
        rows.append((str(entry["sku"]), None))  # Convert to string

# Create a DataFrame
df = pd.DataFrame(rows, columns=["Skus", "vehicle_ids"])
df["concat"] = df["vehicle_ids"] + "&" + df["Skus"]

# Save the DataFrame to a CSV file
df.to_csv(output_file_path, index=False)

print(f"Saved DataFrame to {output_file_path}")
