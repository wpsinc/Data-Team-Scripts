import pandas as pd
import json

# Specify the file path
file_path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\Fitment Json\Online Fitment.json"

# Read the JSON data
with open(file_path, "r") as json_file:
    data = json.load(json_file)

# Extract relevant information from the "vehicles" section (if available)
vehicle_data = data["data"]
rows = []
# Iterate through vehicle data entries
for entry in vehicle_data:
    if "vehicles" in entry and "data" in entry["vehicles"] and entry["vehicles"]["data"]:
        # Assuming you want all vehicle IDs next to the corresponding SKU
        vehicle_ids = []
        for v in entry["vehicles"]["data"]:
            try:
                # Attempt to strip if v["id"] is a string
                vehicle_id = str(v["id"]).strip()
                if vehicle_id != "-":
                    vehicle_ids.append(vehicle_id)
            except AttributeError:
                # If v["id"] is not a string, handle it accordingly
                pass

        sku = str(entry["sku"])  # Convert to string
        for vehicle_id in vehicle_ids:
            rows.append((sku, vehicle_id))
    else:
        # Handle cases where no vehicle data is available
        rows.append((str(entry["sku"]), None))  # Convert to string

# Create a DataFrame
df = pd.DataFrame(rows, columns=["Skus", "vehicle_ids"])
df["concat"] = df["vehicle_ids"] + "&" + df["Skus"]

output_file_path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\onlinefitment.csv"

# Save the DataFrame to a CSV file
df.to_csv(output_file_path, index=False)

print(f"Saved DataFrame to {output_file_path}")
