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

    df[start_col], df[end_col] = zip(
        *df[col_name].apply(
            lambda x: (
                str(x).split("-") if isinstance(x, str) and "-" in x else (str(x), None)
            )
        )
    )

    df[start_col] = df[start_col].str.replace("`", "").str.replace("'", "")
    df[end_col].replace({None: 0}, inplace=True)  # replace None with 0

    df[start_col] = df[start_col].apply(
        lambda x: float(x)
        if isinstance(x, str) and x != "" and pd.notnull(x) and x != "inf"
        else float("nan")
    )
    df[end_col] = df[end_col].apply(
        lambda x: float(x)
        if isinstance(x, str) and x != "" and pd.notnull(x) and x != "inf"
        else float("nan")
    )

    df[start_col] = df[start_col].astype(pd.Int64Dtype())  # convert to nullable integer
    df[end_col] = df[end_col].astype(pd.Int64Dtype())  # convert to nullable integer

    df[start_col] = df[start_col].apply(
        lambda x: 1900 + x if pd.notnull(x) and x >= 50 else 2000 + x
    )
    df[end_col] = df[end_col].apply(
        lambda x: 1900 + x if pd.notnull(x) and x >= 50 else 2000 + x
    )

    df["year_range"] = df.apply(
        lambda row: ",".join(
            [str(i) for i in range(row[start_col], row[end_col] + 1)]
            if pd.notnull(row[start_col])
            and pd.notnull(row[end_col])
            and "-" in str(row[col_name])
            else [str(row[col_name])]
        ),
        axis=1,
    )

    df = df.drop(columns=[start_col, end_col, col_name])
    df = df.rename(columns={"year_range": col_name})

    return df


def reorder_columns(df, original_columns):
    df = df[original_columns]
    return df


def main():
    df = pd.read_excel(file_path, header=None)
    if df.iloc[0].isnull().all():
        df = df.iloc[1:]
    df.reset_index(drop=True, inplace=True)
    df.columns = df.iloc[0]
    df = df.iloc[1:]
    df.reset_index(drop=True, inplace=True)
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
        print("Selected column is empty. Skipping the splitting process.")
    else:
        df = split_year_range(df, col_name)
        df = reorder_columns(df, original_columns)

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
    else:
        print(df.head())


if __name__ == "__main__":
    main()
