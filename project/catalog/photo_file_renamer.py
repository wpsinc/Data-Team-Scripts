import os
import re
import tkinter as tk
from tkinter import filedialog
import inquirer

def rename_files(directory, step, start=0):
    renamed_files = 0
    for filename in os.listdir(directory):
        match = re.match(r"(.*-)(\d+\.?\d*)(_Page_)(\d+)(\.tif)", filename)
        if match:
            base, number, _, page, ext = match.groups()
            if step == 0.1:
                new_number = format(start + (float(page) - 1) * step, '.1f')
            else:
                new_number = str(int(number) + (int(page) - 1) * int(step)).zfill(len(number.split('.')[0]))
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

    questions = [
        inquirer.List('step',
                      message="Select iteration step",
                      choices=['1', '0.1'],
                      ),
    ]
    answers = inquirer.prompt(questions)
    step = float(answers['step'])
    start = 0
    if step == 0.1:
        start = float(input("Enter the starting point: "))
    rename_files(directory, step, start)

select_directory()