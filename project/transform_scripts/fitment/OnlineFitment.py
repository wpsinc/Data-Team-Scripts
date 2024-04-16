import pandas as pd
import json

# Specify the file path
file_path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\Fitment Json\Online Fitment.json"
output_file_path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\onlinefitment.csv"
FitmentFile_Path = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\ToUpload.csv"
output_file_path_Compare = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Fitment Templates\Offroad\onlinefitment_toremove.csv"
Vehicles_file = r"C:\Users\megan.partridge\OneDrive - Arrowhead EP\Data Tech\Fitment Audit\Vehicle File\Vehicles.xlsx"
vehicles_df = pd.read_excel(Vehicles_file)
# Read the JSON data
with open(file_path, "r") as json_file:
    data = json.load(json_file)

# Extract relevant information from the "vehicles" section (if available)
vehicle_data = data["data"]
rows = []

# Iterate through vehicle data entries
for entry in vehicle_data:
    if "vehicles" in entry and "data" in entry["vehicles"] and entry["vehicles"]["data"]:
        # Extract the product ID (ID at the top of each section)
        product_id = entry["id"]
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
            rows.append((sku, product_id, vehicle_id))


# Create a Pandas DataFrame
OnlineFitment = pd.DataFrame(rows, columns=["SKU", "Product_ID", "Vehicle_ID"])

# Print the DataFrame (you can modify this part as needed)

# Create a DataFrame
OnlineFitment = pd.DataFrame(rows, columns=["SKU", "Product_ID" ,"Vehicle_ID"])
OnlineFitment['Vehicle_ID'] = OnlineFitment['Vehicle_ID'].astype(str)
OnlineFitment['SKU'] = OnlineFitment['SKU'].astype(str)
OnlineFitment["concat"] = OnlineFitment["Vehicle_ID"] + "&" + OnlineFitment["SKU"]  # Create the new column

NewFitment = pd.read_csv(FitmentFile_Path, header=None, encoding='utf-8')
NewFitment[0] = NewFitment[0].astype(str)
# Create a new column by concatenating Column1 and Column2
NewFitment["NewColumn"] = NewFitment[0] + "&" + NewFitment[1]
non_matching_rows = OnlineFitment[~OnlineFitment.iloc[:, 3].isin(NewFitment.iloc[:, 2])]
separate_df = non_matching_rows.copy()
separate_df['Product_ID'] = separate_df['Product_ID'].fillna('').astype(str)

# Remove decimal points (if present)
separate_df['Product_ID'] = separate_df['Product_ID'].str.replace(r'\.0$', '', regex=True)

# Create the "URL" column by concatenating the values
separate_df['URL'] = "https://acp.wps-inc.com/items/" + separate_df['Product_ID'] + "/edit"
separate_df['Vehicle_ID'] = pd.to_numeric(separate_df['Vehicle_ID'], errors='coerce')
separate_df = pd.merge(separate_df, vehicles_df, left_on='Vehicle_ID', right_on='vehicle_id')
selected_columns = ['URL', 'SKU', 'vehicle_type', 'year', 'make', 'model']
filtered_df = separate_df [selected_columns]
# Display a preview of the data
# Display data in chunks of 25 rows
#chunk_size = 25
#total_rows = len(filtered_df)

# Initialize an empty list to store approved indices
#approved_indices = []

# for start_idx in range(0, total_rows, chunk_size):
   #  end_idx = min(start_idx + chunk_size, total_rows)
   #  chunk_df = filtered_df.iloc[start_idx:end_idx]

    # print("\nChunk of rows {} to {}:".format(start_idx + 1, end_idx))
    # print(chunk_df)  # Display the chunk of data

    # Prompt user for approval
    # user_approval = input("Is this chunk okay? (y/n): ").lower()
    # if user_approval == 'y':
        # User approves, add indices to the approved list
     #   approved_indices.extend(range(start_idx, end_idx))
    # else:
        # print("Chunk not approved. Moving to the next chunk.")

# Filter the data based on approved indices
# cleaned_df = filtered_df.loc[approved_indices][['URL']].drop_duplicates()

# Save the cleaned data to a new file (replace 'cleaned_data.csv' with your desired filename)
# cleaned_df.to_csv('cleaned_data.csv', index=True)

# print("Cleaned data saved to 'onlinefitment_toremove.csv'.")

filtered_df.to_csv(output_file_path_Compare)
# Save the DataFrame to a CSV file
OnlineFitment.to_csv(output_file_path, index=False) 

print(f"Saved DataFrame to {output_file_path}")
print(f"Saved DataFrame to {output_file_path_Compare}")
