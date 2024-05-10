"""
The script first imports necessary modules: os for interacting with the operating system, re for regular expressions, and tkinter for creating a graphical user interface (GUI).

The rename_files function is defined. This function takes a directory path and an optional step argument (defaulting to 1). It then iterates over each file in the provided directory.

For each file, it uses a regular expression to match filenames that follow a specific pattern: (.*)(_Page_)(\d+)(\.tif). This pattern matches any filename that ends with "Page" followed by one or more digits and a ".tif" extension.

If a filename matches this pattern, it is split into components: the base name, the page number, and the extension. The page number is then zero-padded to 4 digits.

A new filename is constructed using the base name, the zero-padded page number, and the extension. The original file is then renamed to this new filename.

If no files were renamed in the directory, a message is printed to the console.

The select_directory function is defined. This function creates a simple GUI that allows the user to select a directory. Once a directory is selected, the rename_files function is called on that directory.

Finally, the select_directory function is called to start the process.
"""

import os
import re
import tkinter as tk
from tkinter import filedialog

def rename_files(directory, step=1):
    renamed_files = 0
    for filename in os.listdir(directory):
        match = re.match(r"(.*)(_Page_)(\d+)(\.tif)", filename)
        if match:
            base, _, page, ext = match.groups()
            new_number = str(int(page)).zfill(4)  # zero-fill to 4 digits
            new_filename = base + new_number + ext
            print(f"Old File Name: {filename} Changed to: {new_filename}")
            os.rename(os.path.join(directory, filename), os.path.join(directory, new_filename))
            renamed_files += 1
    if renamed_files == 0:
        print("No files were renamed.")

def select_directory():
    root = tk.Tk()
    root.withdraw()
    directory = filedialog.askdirectory(title="Select Folder to Rename Files")
    rename_files(directory)

select_directory()