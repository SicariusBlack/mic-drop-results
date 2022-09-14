import json
import os
import pandas as pd
import pathlib
import re
import requests
import subprocess

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.slide import Slide

import win32com
import win32com.client


def hex_to_rgb(hex):
    hex = hex.lstrip("#")
    return tuple(int(hex[i:i + 2], 16) for i in (0, 2, 4))


def replace_text(slide: Slide, df, search_str: str, repl: str) -> Slide:
    """Replaces and formats text
    
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


# Section B: Check for Updates
status = ""

if config["update_check"]:
    try:
        response = requests.get("https://api.github.com/repos/"
            "berkeleyfx/mic-drop-results/releases/latest", timeout=3)

        version = float(response.json()["tag_name"][1:])
        
        if version > config["version"]:
            print(f"A new version (v{version}) is available. "
                "You can download it using the link below.")
            print("https://github.com/berkeleyfx/mic-drop-results/releases/latest/\n")

            status = "update available"
        elif version < config["version"]:
            status = "beta"
        else:
            status = "up to date"
        
        status = " [" + status + "]"
    except:
        pass

print(f"Mic Drop Results (v{config['version']}){status}")

if not "update available" in status:
    print("https://github.com/berkeleyfx/mic-drop-results")


# Section C: Data Cleaning
df = pd.read_excel("data.xlsx")

# Check for cases where avg and std are the same (hold the same rank)
df["r"] = pd.DataFrame(zip(df.iloc[:, 0], df.iloc[:, 1] * -1)) \
    .apply(tuple, axis=1).rank(method="min", ascending=False).astype(int)

# Sort the slides
df = df.sort_values(by="r", ascending=True)

# Remove .0 from whole numbers
format_number = lambda x: str(int(x)) if x % 1 == 0 else str(x)
df.loc[:, df.dtypes == float] = df.loc[:, df.dtypes == float].applymap(format_number)


# Section D: To PowerPoint
print("\nGenerating slides...")
print("Please do not click on any PowerPoint windows that may show up in the process.")
print("Try hitting Enter if the program does not respond for more than 15 seconds.")

# Kill all PowerPoint instances
subprocess.run("TASKKILL /F /IM powerpnt.exe",
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

# Open template presentation
path = str(pathlib.Path().resolve()) + "\\"

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Presentations.Open(f"{path}template.pptm")

# Import module
try:
    ppt.VBE.ActiveVBProject.VBComponents.Import(f"{path}Module1.bas")
except:
    input("\nERROR: Please open PowerPoint, look up Trust Center Settings, "
        "and make sure Trust access to the VBA project object model is checked.")

# Running VBA Functions
slides_count = ppt.Run("Count")

for t in df.loc[:, "template"]:
    ppt.Run("Duplicate", t)

# Delete template slides when done
ppt.Run("DelSlide", *range(1, slides_count + 1))

# Save as output file
output_filename = "output.pptx"
path += "output\\"

pathlib.Path(path).mkdir(parents=True, exist_ok=True)

ppt.Run("SaveAs", f"{path}{output_filename}")
ppt.Quit()


# Section E: Fill in the Blank
prs = Presentation(path + output_filename)

for i, slide in enumerate(prs.slides):
    for col in df.columns:
        replace_text(slide, df, "{" + col + "}", str(df[col].iloc[i]))

prs.save(path + output_filename)


# Section F: Launching the File
print(f"\nExported to {path}{output_filename}")
input("Press Enter to launch the file...")

os.startfile(path + output_filename)
