import json
import os
import pandas as pd
import pathlib
import re
import requests
import subprocess

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.slide import Slide

import win32com
import win32com.client


def hex_to_rgb(hex):
    hex = hex.lstrip("#")
    return tuple(int(hex[i:i + 2], 16) for i in (0, 2, 4))


def replace_text(slide: Slide, df, i) -> Slide:
    """Replaces and formats text"""
    cols = df.columns.values.tolist()

    for shape in slide.shapes:
        if not shape.has_text_frame or not "{" in shape.text:
            continue

        text_frame = shape.text_frame

        for run in [p.runs[0] for p in text_frame.paragraphs]:
            for search_str in set(re.findall(r"(?<={)(.*?)(?=})", run.text)).intersection(cols):
                repl = str(df[search_str].iloc[i])
                run.text = run.text.replace("{" + search_str + "}", repl)

                if not search_str.startswith(starts) or not run.font.color.type:
                    continue

                if run.font.color.type == MSO_COLOR_TYPE.RGB:
                    if not run.font.color.rgb == RGBColor(255, 255, 255):
                        continue

                for ind, val in enumerate(range_list):
                    if float(repl) >= val:
                        run.font.color.rgb = RGBColor(*col_list[ind])
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
        pass  # Move on if there is no internet connection

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
print("Try hitting Enter if the program does not respond for more than 10 seconds.")

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

os.makedirs(path, exist_ok=True)

ppt.Run("SaveAs", f"{path}{output_filename}")
ppt.Quit()


# Section E: Fill in the Blank
prs = Presentation(path + output_filename)

for i, slide in enumerate(prs.slides):
    replace_text(slide, df, i)

prs.save(path + output_filename)


# Section F: Launching the File
print(f"\nExported to {path}{output_filename}")
input("Press Enter to launch the file...")

os.startfile(path + output_filename)
