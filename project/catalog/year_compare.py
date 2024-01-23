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


def compare_dataframes(df1, df2):
    df1.set_index("WICITEM", inplace=True)
    df2.set_index("WICITEM", inplace=True)

    df2["isNew"] = ~df2.index.isin(df1.index)

    df1, df2 = df1.align(df2, axis=1)
    df1, df2 = df1.align(df2, axis=0)

    comparison_df = df1.compare(df2)

    comparison_df.columns = comparison_df.columns.set_levels(["2023", "2024"], level=1)

    return comparison_df


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

        comparison_df = compare_dataframes(old_year_df, new_year_df)

        # Get the name of the last column
        column_name = comparison_df.columns[-1]

        # Fill blank values in the column with 'REMOVED'
        comparison_df.loc[comparison_df[column_name].isna(), column_name] = "REMOVED"

        output_file = os.path.join(os.path.dirname(new_year_file), "changes.xlsx")
        comparison_df.to_excel(output_file)


if __name__ == "__main__":
    main()
