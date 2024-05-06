"""
This Python script is designed to compare two Excel files (specifically, .xlsx files) from different years and identify the changes between them. The script uses the pandas library to handle the Excel files as dataframes, and the tkinter library to open a file dialog for the user to select the files to compare. Here's a step-by-step breakdown:

The select_files function opens a file dialog for the user to select multiple Excel files. It then filters the selected files to only include those with "Final.xlsx" in their names, sorts them, and returns the sorted list of file paths.

The compare_dataframes function takes two dataframes and two file names as input. It sets the index of both dataframes to the "WICITEM" column, then adds a new column to the second dataframe that indicates whether each row is new (i.e., not present in the first dataframe). It aligns the two dataframes along both axes, then creates a new dataframe that compares the two original dataframes. It also extracts the first four characters from each file name (presumably the year), and uses these to set the levels of the comparison dataframe's columns. It returns the comparison dataframe along with the extracted years.

The main function is the entry point of the script. It prompts the user to select the year files, then reads the first two selected files into dataframes. It uses the compare_dataframes function to compare these dataframes and get the comparison dataframe and the years. It then fills any blank values in the last column of the comparison dataframe with the string "REMOVED". Finally, it saves the comparison dataframe to an Excel file in the user's Downloads folder, with a filename formatted using the years.

If the script is run as a standalone program (i.e., not imported as a module), it calls the main function to start the comparison process.
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from rich.progress import Progress

def select_files():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])

    file_paths = [file_path for file_path in file_paths if "Final.xlsx" in file_path]

    file_paths = sorted(file_paths)

    return file_paths

def compare_dataframes(df1, df2, old_year_file, new_year_file):
    df1.set_index("WICITEM", inplace=True)
    df2.set_index("WICITEM", inplace=True)

    df2["isNew"] = ~df2.index.isin(df1.index)

    df1, df2 = df1.align(df2, axis=1)
    df1, df2 = df1.align(df2, axis=0)

    comparison_df = df1.compare(df2)

    # Extract the first 4 characters from the filenames
    old_year = os.path.basename(old_year_file)[:4]
    new_year = os.path.basename(new_year_file)[:4]

    # Use the extracted years to set the levels of the columns
    comparison_df.columns = comparison_df.columns.set_levels([old_year, new_year], level=1)

    # Return the extracted years along with the comparison DataFrame
    return comparison_df, old_year, new_year

def main():
    print("Select the year files")
    year_files = select_files()

    old_year_file = year_files[0]
    new_year_file = year_files[1]

    with Progress() as progress:
        task = progress.add_task("[cyan]Processing...", total=100)

        old_year_df = pd.read_excel(old_year_file)
        progress.update(task, advance=50)

        new_year_df = pd.read_excel(new_year_file)
        progress.update(task, advance=50)

        comparison_df, old_year, new_year = compare_dataframes(old_year_df, new_year_df, old_year_file, new_year_file)

        # Get the name of the last column
        column_name = comparison_df.columns[-1]

        # Fill blank values in the column with 'REMOVED'
        comparison_df.loc[comparison_df[column_name].isna(), column_name] = "REMOVED"

        # Get the directory of the current script
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

        # Format the output filename using the extracted years
        output_file = os.path.join(downloads_folder, f"{old_year}to{new_year}_Catalog_Changes.xlsx")
        comparison_df.to_excel(output_file)

if __name__ == "__main__":
    main()