import pandas as pd
import os
import shutil
# Get the current user's home directory
home_dir = os.path.expanduser("~")

# Replace hardcoded paths
input_folder = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/Item Maintenance/Source/Lists to Append"
)
output_folder = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/Item Maintenance/Source"
)
output_file = os.path.join(output_folder, "ItemMaintLists.csv")
common_end = '.csv'
dataframes = []
for file in os.listdir(input_folder):
    if file.endswith(common_end):
        file_path = os.path.join(input_folder, file)
        read_file = pd.read_csv(file_path)
        dataframes.append(read_file)

combined_read_files = pd.concat(dataframes, ignore_index=True)

combined_read_files.to_csv(output_file, index=False)

for filename in os.listdir(input_folder):
    file_path = os.path.join(input_folder, filename)
    try:
        if os.path.isfile(file_path) or os.path.islink(file_path):
            os.unlink(file_path)  # Remove the file
        elif os.path.isdir(file_path):
            os.rmdir(file_path)  # Remove the directory
    except Exception as e:
        print(f'Failed to delete {file_path}. Reason: {e}')
shutil.move(output_file, input_folder)