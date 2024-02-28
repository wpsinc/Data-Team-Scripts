import os
import json
from PIL import Image
from tkinter import filedialog
from tkinter import Tk

root = Tk()
root.withdraw()
folder_path = filedialog.askdirectory()

try:
    with open("lessThan300dpi.json", "r") as f:
        data = json.load(f)
        existing_paths = {item["path"] for item in data}
except (FileNotFoundError, json.JSONDecodeError):
    data = []
    existing_paths = set()

for subdir, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith(".jpg") or file.endswith(".png") or file.endswith(".tif"):
            img_path = os.path.join(subdir, file)
            if img_path in existing_paths:
                continue
            img = Image.open(img_path)
            dpi = img.info.get("dpi")
            if dpi is not None and dpi[0] < 300:
                print(f"Less than 300dpi: {img_path}")
                data.append(
                    {
                        "path": img_path,
                        "dpi": {"horizontal": dpi[0], "vertical": dpi[1]},
                    }
                )

with open("lessThan300dpi.json", "w") as f:
    json.dump(data, f, indent=4)