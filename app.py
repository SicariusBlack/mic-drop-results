import json
import numpy as np
import os
import pandas as pd
import pathlib
import re

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.slide import Slide

import win32com
import win32com.client


def hex_to_rgb(hex):
    hex = hex.lstrip("#")
    return tuple(int(hex[i:i + 2], 16) for i in (0, 2, 4))


def replace_text(slide: Slide, search_str: str, repl: str) -> Slide:
    """
    Modified function from the pptx_replace package
    https://github.com/PaleNeutron/pptx-replace/blob/master/pptx_replace/replace_core.py
    """
    search_pattern = re.compile(re.escape(search_str), re.IGNORECASE)

    for shape in slide.shapes:
        if shape.has_text_frame and re.search(search_pattern, shape.text):
            text_frame = shape.text_frame

            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    if re.search(search_pattern, run.text):
                        run.text = re.sub(search_pattern, repl, run.text)

                    if search_str[1:].startswith(starts) and run.font.color.type:
                        for i, val in enumerate(range_list):
                            if float(repl) >= val:
                                run.font.color.rgb = RGBColor(*col_list[i])
                                break
    return slide


# Section A: Load config.json
config = json.load(open("config.json"))

# Variable shortcuts
range_list = config["format"]["ranges"][::-1]
col_list = config["format"]["colors"][::-1]
starts = config["format"]["starts_with"]

col_list = list(map(hex_to_rgb, col_list))

# Section B: Data Cleaning
print(f"Mic Drop Results (Version {config['version']})")
print("https://github.com/berkeleyfx/mic-drop-results")

df = pd.read_excel("data.xlsx")

# Check for cases where avg and std are the same (hold the same rank)
df["r"] = pd.DataFrame(zip(df.iloc[:, 0], df.iloc[:, 1] * -1)) \
    .apply(tuple, axis=1).rank(method="min", ascending=False).astype(int)

# Sort the slides
df = df.sort_values(by="r", ascending=True)

# Remove .0 from whole numbers
format_number = lambda x: str(int(x)) if x % 1 == 0 else str(x)
df.loc[:, df.dtypes == float] = df.loc[:, df.dtypes == float].applymap(format_number)


# Section C: To PowerPoint
print("\nGenerating slides...")
print("Please do not click on any PowerPoint windows that may show up in the process.")
print("Try hitting Enter if the program freezes for more than 30 seconds.")

path = str(pathlib.Path().resolve()) + "\\"

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Presentations.Open(f"{path}template.pptm")

# Exported Module
ppt.VBE.ActiveVBProject.VBComponents.Import(f"{path}Module1.bas")

# Running VBA Functions
slides_count = ppt.Run("Count")

for t in df.loc[:, "template"]:
    ppt.Run("Duplicate", t)

# Delete template slides when done
for i in range(slides_count):
    ppt.Run("DelSlide", 1)

output_filename = "output.pptx"
path += "output\\"

pathlib.Path(path).mkdir(parents=True, exist_ok=True)

ppt.Run("SaveAs", f"{path}{output_filename}")
ppt.Quit()


# Section D: Fill in the Blank
prs = Presentation(path + output_filename)

for i, slide in enumerate(prs.slides):
    for col in df.columns:
        replace_text(slide, "{" + col + "}", str(df[col].iloc[i]))

prs.save(path + output_filename)


# Section D: Launching the File
print(f"\nExported to {path}{output_filename}")
input("Press Enter to launch the file...")

os.startfile(path + output_filename)
