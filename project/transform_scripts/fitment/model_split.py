import pandas as pd
from tkinter import filedialog
from tkinter import Tk
import inquirer
from tqdm import tqdm
from colorama import Fore, Style

root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title="Select File to Split the Text Column on")


def split_text(df, col_name, original_columns):
    new_rows = []
    for row in df.itertuples(index=False):
        text = getattr(row, col_name)
        if "/" in text:
            split_texts = text.split("/")
            for split_text in split_texts:
                new_row = list(row)
                new_row[original_columns.index(col_name)] = split_text.strip()
                new_rows.append(new_row)
        else:
            new_rows.append(list(row))
    new_df = pd.DataFrame(new_rows, columns=df.columns)
    return new_df


def main():
    print(Fore.GREEN + "Welcome to the text splitter script!" + Style.RESET_ALL)

    df = pd.read_excel(file_path, header=None)
    if df.loc[0].isnull().all():
        df = df.loc[1:]
    df.reset_index(drop=True, inplace=True)
    df.columns = df.loc[0]
    df = df.loc[1:]

    original_columns = df.columns.tolist()
    questions = [
        inquirer.List(
            "col_name",
            message="Enter the name of the column to split",
            choices=original_columns,
        ),
    ]
    answers = inquirer.prompt(questions)
    col_name = answers["col_name"]

    if df[col_name].isnull().all():
        print(Fore.RED + "Selected column is empty. Skipping the splitting process." + Style.RESET_ALL)
    else:
        print("Splitting text...")
        df = split_text(df, col_name, original_columns)  # Pass original_columns here
        print(Fore.GREEN + "Splitting completed!" + Style.RESET_ALL)

    update_choices = [
        inquirer.List(
            "update_excel",
            message="Do you want to update the Excel file selected?",
            choices=["Yes", "No, just testing"],
        ),
    ]
    update_answer = inquirer.prompt(update_choices)
    if update_answer["update_excel"] == "Yes":
        df.to_excel(file_path, index=False)
        print(Fore.GREEN + "Excel file updated successfully!" + Style.RESET_ALL)
    else:
        print(df.head())

    print(Fore.GREEN + "Thank you for using the text splitter script!" + Style.RESET_ALL)

if __name__ == "__main__":
    main()
