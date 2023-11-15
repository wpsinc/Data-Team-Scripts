import pandas as pd
from tkinter import filedialog
from tkinter import Tk
import inquirer

root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title="Select File to Split the Year Column on")


def split_year_range(df, col_name):
    start_col = "start_year"
    end_col = "end_year"
    df[[start_col, end_col]] = df[col_name].str.split("-", expand=True)

    df[start_col] = df[start_col].str.replace("`", "")

    df[start_col] = df[start_col].astype(int)
    df[end_col] = df[end_col].astype(int)

    df[start_col] = df[start_col].apply(lambda x: 1900 + x if x >= 50 else 2000 + x)
    df[end_col] = df[end_col].apply(lambda x: 1900 + x if x >= 50 else 2000 + x)

    df["year_range"] = df.apply(
        lambda row: ",".join(
            [str(i) for i in range(row[start_col], row[end_col] + 1)]
            if "None" not in str(row[start_col]) and "None" not in str(row[end_col])
            else ""
        ),
        axis=1
    )

    df = df.drop(columns=[start_col, end_col, col_name])
    df = df.rename(columns={"year_range": col_name})

    return df

# def reorder_columns(df):
#     df = df[
#
#     ]
#     return df

def main():
    df = pd.read_excel(file_path, header=None)
    df = df.iloc[1:]
    df.reset_index(drop=True, inplace=True)
    df.columns = df.iloc[0]
    df = df.iloc[1:] 
    df.reset_index(drop=True, inplace=True)
    columns = df.columns.tolist()
    questions = [
        inquirer.List('col_name',
                      message="Enter the name of the column to split",
                      choices=columns),
    ]
    answers = inquirer.prompt(questions)
    col_name = answers['col_name']
    df = split_year_range(df, col_name)
    # df = reorder_columns(df)
    print(df.head())
    # df.to_excel(file_path, index=False)


if __name__ == "__main__":
    main()
