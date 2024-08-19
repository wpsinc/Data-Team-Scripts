import pandas as pd
import json
import os
# Get the current user's home directory
home_dir = os.path.expanduser("~")
# Specify the file path
def display_menu(options):
    """
    Displays a menu with the provided options.
    Args:
        options (list): List of menu options.
    """
    print("Select an option:")
    for i, option in enumerate(options, start=1):
        print(f"{i}. {option}")

def get_user_choice(options):
    """
    Gets the user's choice from the menu.
    Args:
        options (list): List of menu options.
    Returns:
        int: User's selected option (1-based index).
    """
    while True:
        try:
            choice = int(input("Enter the number of your choice: "))
            if 1 <= choice <= len(options):
                return choice
            else:
                print("Invalid choice. Please select a valid option.")
        except ValueError:
            print("Invalid input. Please enter a number.")

# List of categories
categories = ["Offroad", "ATV", "Snow", "Street", "Watercraft"]

# Display the menu
display_menu(categories)

# Get user's choice
selected_index = get_user_choice(categories)
selected_category = categories[selected_index - 1]

# Construct the paths
file_path = os.path.join(
    home_dir, f"OneDrive - Arrowhead EP/Data Tech/Fitment Audit/Fitment Templates/{selected_category}/Fitment Json/Online Fitment.json"
)
output_file_path = os.path.join(
    home_dir, f"OneDrive - Arrowhead EP/Data Tech/Fitment Audit/Fitment Templates/{selected_category}/onlinefitment.csv"
)
FitmentFile_Path = os.path.join(
    home_dir, f"OneDrive - Arrowhead EP/Data Tech/Fitment Audit/Fitment Templates/{selected_category}/ToUpload.csv"
)
output_file_path_Compare = os.path.join(
    home_dir, f"OneDrive - Arrowhead EP/Data Tech/Fitment Audit/Fitment Templates/{selected_category}/onlinefitment_toremove.csv"
)
Vehicles_file = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/Fitment Audit/Vehicle File/Vehicles.xlsx"
)
output_path = os.path.join(
    home_dir, f"OneDrive - Arrowhead EP/Data Tech/Fitment Audit/Fitment Removed and Updated"
)
NewFitment = pd.read_csv(FitmentFile_Path, header=None, encoding='utf-8')
# Copy the 2nd Column into a Separate DataFrame
Third_column_df = NewFitment.iloc[:, 2].copy()

# Remove Duplicates from the Second Column
Third_column_df.drop_duplicates(inplace=True)

# Concatenate the Column into a Single String
concatenated_string = ','.join(Third_column_df.astype(str))

# Print the Concatenated String for the User
print(f"Concatenated string from the 2nd column:\n{concatenated_string}")

# Wait for user confirmation
input("Press Enter when you're ready to continue...")
# Replace NaN values with 0
NewFitment[1] = NewFitment[1].fillna(0)

# Now convert to integer and then to string
NewFitment[1] = NewFitment[1].astype(int).astype(str)
grouping_column = 'grouping_numbers'
# Group by the first column
grouped = NewFitment.groupby(NewFitment.columns[0])

# Create separate DataFrames for each group
for group_number, group_df in grouped:
    # Drop the first column (index 0)
    group_df = group_df.iloc[:, 1:]
    
    # Create the filename (e.g., '1-Fitment.csv')
    filename = f"{group_number}- Fitment.csv"
    
    # Write the group DataFrame to a CSV file without index columns
    output_csv = os.path.join(output_path, filename)
    group_df.to_csv(output_csv, index=False, header=False)  # Avoid writing index columns
    print(f"Saved {filename} to {output_csv}")

print("All groups saved to separate CSV files.")
# Create a new column by concatenating Column1 and Column2
NewFitment["NewColumn"] = NewFitment[1] + "&" + NewFitment[2]
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

non_matching_rows = OnlineFitment[~OnlineFitment.iloc[:, 3].isin(NewFitment.iloc[:, 3])]
separate_df = non_matching_rows.copy()
separate_df['Product_ID'] = separate_df['Product_ID'].fillna('').astype(str)

# Remove decimal points (if present)
separate_df['Product_ID'] = separate_df['Product_ID'].str.replace(r'\.0$', '', regex=True)

# Create the "URL" column by concatenating the values
separate_df['URL'] = "https://acp.wps-inc.com/items/" + separate_df['Product_ID'] + "/edit"
separate_df['Vehicle_ID'] = pd.to_numeric(separate_df['Vehicle_ID'], errors='coerce')
separate_df = pd.merge(separate_df, vehicles_df, left_on='Vehicle_ID', right_on='vehicle_id')
selected_columns = ['Vehicle_ID', 'URL', 'SKU', 'vehicle_type', 'year', 'make', 'model']
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