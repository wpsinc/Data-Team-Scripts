import pandas as pd
import os
import time
from halo import Halo

# Get the current user's home directory
home_dir = os.path.expanduser("~")

# Define paths
output_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/End of Month Templates/Linked Reports/Bobs Mega Report/Output"
)
output_folder = os.path.join(output_path, "Bobs Mega Report Output")
output_file = os.path.join(output_path, "Bobs Merged Mega Report Output.csv")

# Check if the output file exists
if not os.path.exists(output_file):
    print(f"Error: The output file {output_file} does not exist.")
    exit()

# Ensure the output folder exists (it will be created if it doesn't)
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Start loading indicator
spinner = Halo(text="Reading merged output file...", spinner="dots")
spinner.start()
start_time = time.time()

# Read the merged output file
df = pd.read_csv(output_file, encoding='utf-8', low_memory=False)

end_time = time.time()
spinner.stop_and_persist(
    symbol="✔️ ".encode("utf-8"),
    text=f"File read successfully. Process Time: {round((end_time - start_time)/60, 5)} minutes",
)

# Check if 'Invoice Year' column exists
if "Invoice Year" in df.columns:
    spinner = Halo(text="Writing data to Excel files by year...", spinner="dots")
    spinner.start()
    start_time = time.time()

    unique_years = sorted(df["Invoice Year"].dropna().unique())

    # Loop through each year
    for year in unique_years:
        year_data = df[df["Invoice Year"] == year]
        rows_written = 0
        file_suffix = 1
        year_file_path = os.path.join(output_folder, f"Bobs_Mega_Report_{year}.xlsx")

        with pd.ExcelWriter(year_file_path, engine='openpyxl') as writer:
            # Split data into chunks of 800,000 rows
            for chunk in range(0, len(year_data), 800000):
                chunk_data = year_data.iloc[chunk:chunk + 800000]

                # Write the chunk to the current sheet, ensuring the header row is included
                chunk_data.to_excel(writer, sheet_name=f"Sheet{file_suffix}", index=False)

                rows_written += len(chunk_data)
                file_suffix += 1

                print(f"✔️ Successfully wrote data for year {year} to sheet 'Sheet{file_suffix - 1}'")

    end_time = time.time()
    spinner.stop_and_persist(
        symbol="✔️ ".encode("utf-8"),
        text=f"Data successfully written to {output_folder}. Process Time: {round((end_time - start_time)/60, 5)} minutes",
    )
else:
    print("Warning: 'Invoice Year' column not found in the DataFrame.")

end_time = time.time()
print(f"Total process took {round((end_time - start_time)/60, 5)} minutes")
