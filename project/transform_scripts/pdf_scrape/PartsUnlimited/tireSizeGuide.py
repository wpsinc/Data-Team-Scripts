import os
import re
import uuid
from tkinter import Tk, filedialog
from PyPDF2 import PdfReader
import pandas as pd

root = Tk()
root.withdraw()
files = filedialog.askopenfilenames(title="Select PDF(s) to Read")


def process_file(file):
    pdf = PdfReader(file)

    text = ""
    for page in pdf.pages:
        text += page.extract_text()

    def add_amp(match):
        year_range = match.group(0)
        numbers = list(map(int, year_range.split("-")))
        if any(25 <= number <= 40 for number in numbers):
            return year_range
        return f"&{year_range}&"

    text = re.sub(r"(?<!/)\b\d{2}-\d{2}\b", add_amp, text)
    text = re.sub(r"(?<= )\b\d{2}\b(?= )", add_amp, text)
    text = text.split("FITS MODEL FRONT TIRE REAR TIRE", 1)[1]

    data = []
    for line in text.split("\n"):
        text = text.strip()
        values = line.split("&")
        data.append(values)

    df = pd.DataFrame(
        data,
        columns=["Model", "Year", "Front Tire", "Rear Tire", "Additional Column(s)"],
    )

    base_name = os.path.splitext(os.path.basename(file))[0]
    file_id = uuid.uuid4()

    output_path = f"C:/Users/London.Perry/Downloads/{base_name}_{file_id}.xlsx"
    df.to_excel(output_path, index=False)

    print("Success! File saved at:", output_path)

    directory = os.path.dirname(output_path)
    os.startfile(directory)


for file in files:
    process_file(file)
