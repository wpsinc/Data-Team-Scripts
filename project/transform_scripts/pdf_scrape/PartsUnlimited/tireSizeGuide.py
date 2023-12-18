import os
import re
import uuid
import io
from tkinter import Tk, filedialog
from PyPDF2 import PdfReader
import pandas as pd

root = Tk()
root.withdraw()
files = filedialog.askopenfilenames(title="PDFs to Read")

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

    df = pd.read_csv(
        io.StringIO(text), names=["Model", "Year", "Front Tire", "Rear Tire"]
    )

    base_name = os.path.splitext(os.path.basename(file))[0]
    file_id = uuid.uuid4()

    df.to_excel(
        f"C:/Users/London.Perry/Downloads/{base_name}_{file_id}.xlsx", index=False
    )

for file in files:
    process_file(file)