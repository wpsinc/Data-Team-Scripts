import pandas as pd
import json

# Specify the file path for the JSON data
products = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\Fitment Json\ItemGroups.json"

# Read the JSON data
with open(products, "r") as json_file:
    data = json.load(json_file)

numeric_ids_set = set()
for product in data:  # Iterate over the loaded data, not 'products'
    for group in product["cataloggroups"]["data"]:
        numeric_ids_set.append(group["id"])
df = pd.DataFrame({
    "id": numeric_ids_set
})



print(df)