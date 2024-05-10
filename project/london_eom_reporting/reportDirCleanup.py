"""
This Python script is designed to organize report files in a directory structure based on the year and month information extracted from the file names. The script uses the os and shutil libraries to interact with the file system, and the tkinter library to open a directory selection dialog. Here's a step-by-step breakdown:

The main function is the entry point of the script. It first opens a directory selection dialog for the user to select the directory containing the reports.

It then lists all the subdirectories in the selected directory.

For each subdirectory, it lists all the files in it. If there are no files, it prints a message and moves on to the next subdirectory.

For each file, it tries to split the file name on the underscore character and then on the dash character to extract the year and month. If this fails (i.e., if the file name doesn't contain these characters), it gets the file's creation date and uses that as the year and month. It also renames the file to include this date.

It then checks if a directory for the year exists in a "Report History" subdirectory of the original subdirectory. If not, it creates this directory.

It then checks if a directory for the month exists in the year directory. If not, it creates this directory.

Finally, it tries to move the file to the month directory. If this fails, it prints an error message.

If the script is run as a standalone program (i.e., not imported as a module), it calls the main function to start the organization process.

The script is designed to be run periodically to keep the report files organized. It assumes that the file names follow a specific format, and that the directory structure is set up in a specific way.
"""

import os
import shutil
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askdirectory

def main():
    Tk().withdraw() 
    reports_dir = askdirectory(title="Select the directory containing the reports")  
    print(f"reports_dir: {reports_dir}")
    vendor_dirs = [
        d for d in os.listdir(reports_dir) if os.path.isdir(os.path.join(reports_dir, d))
    ]

    for vendor_dir in vendor_dirs:
        src_dir = os.path.join(reports_dir, vendor_dir)
        dst_dir = os.path.join(src_dir, "Report History")

        files = [f for f in os.listdir(src_dir) if os.path.isfile(os.path.join(src_dir, f))]

        if not files:
            dir_name = os.path.basename(src_dir)
            print(f"No files to process for {dir_name}")
            continue

        for file in files:
            try:
                year, month, _ = file.split("_")[1].split("-")
            except IndexError:
                creation_date = datetime.fromtimestamp(
                    os.path.getctime(os.path.join(src_dir, file))
                ).strftime("%Y-%m")
                actual_dir_name = os.path.basename(os.path.dirname(os.path.join(src_dir, file)))
                new_file = f"{actual_dir_name}_{creation_date}-01.{file.split('.')[1]}"
                os.rename(os.path.join(src_dir, file), os.path.join(src_dir, new_file))
                file = new_file
                year, month = creation_date.split("-")

            year_dir = os.path.join(dst_dir, year)

            if not os.path.exists(year_dir):
                os.makedirs(year_dir)
                print(f"Created new directory: {year_dir}")

            month_dir = os.path.join(year_dir, f"{month}-{year[2:]}")

            if not os.path.exists(month_dir):
                os.makedirs(month_dir)
                print(f"Created new directory: {month_dir}")

            try:
                shutil.move(os.path.join(src_dir, file), os.path.join(month_dir, file))
                print(f"Successfully moved {file} to {actual_dir_name}_{month}-{year[2:]}")
            except Exception as e:
                print(f"Error moving {file} to {month_dir}: {e}")

if __name__ == '__main__':
    main()