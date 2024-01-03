import os
import shutil
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()
reports_dir = os.getenv("REPORTS_DIR")
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
            print(f"Successfully moved {file} to {dir_name}_{month}-{year[2:]}")
        except Exception as e:
            print(f"Error moving {file} to {month_dir}: {e}")
