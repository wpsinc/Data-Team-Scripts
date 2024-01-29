import os
import re
import tkinter as tk
from tkinter import filedialog

def rename_files(directory):
    for filename in os.listdir(directory):
        match = re.match(r"(.*-)(\d+)(_Page_)(\d+)(\.tif)", filename)
        if match:
            base, number, _, page, ext = match.groups()
            new_number = str(int(number) + int(page) - 1).zfill(len(number))
            new_filename = base + new_number + ext
            print(f"Old File Name: {filename} Changed to: {new_filename}")
            os.rename(os.path.join(directory, filename), os.path.join(directory, new_filename))

def select_directory():
    root = tk.Tk()
    root.withdraw()
    directory = filedialog.askdirectory(title="Select Folder to Rename Files")
    rename_files(directory)

select_directory()