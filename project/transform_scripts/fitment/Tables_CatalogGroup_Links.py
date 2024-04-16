import pandas as pd
import json

# Specify the file path for the JSON data
file_path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\Fitment Json\Tables.json"

# Read the JSON data
with open(file_path, "r") as json_file:
    data = json.load(json_file)

# Initialize empty lists to store extracted values
ids = []
cataloggroup_ids = []
urls = []

# Extract id and cataloggroup_id for each entry
for entry in data["data"]:
    ids.append(entry["id"])
    cataloggroup_ids.append(entry["cataloggroup_id"])
    urls.append(f"https://acp.wps-inc.com/cataloggroups/{entry['cataloggroup_id']}/edit")

# Create a DataFrame with the extracted data
df = pd.DataFrame({
    "id": ids,
    "cataloggroup_id": cataloggroup_ids,
    "URL": urls
})

# Specify the output CSV file path
output_csv_path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\CatalogGroupsLinks.csv"

# Write the DataFrame to CSV
df.to_csv(output_csv_path, index=False)

print(f"Data written to {output_csv_path}")
