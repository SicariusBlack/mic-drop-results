import ctypes
import json
import os
import pandas as pd
import pathlib
import re
import requests
import signal
import subprocess

from alive_progress import alive_bar

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


def link(uri, label=None):
    """Prints text as a clickable link."""
    if label is None: 
        label = uri
    parameters = ''

    # OSC 8 ; params ; URI ST <name> OSC 8 ;; ST 
    escape_mask = '\033]8;{};{}\033\\{}\033]8;;\033\\'

    return escape_mask.format(parameters, uri, label)


# Section A: Loading config.json
config = json.load(open("config.json"))

# Variable shortcuts
range_list = config["format"]["ranges"][::-1]
col_list = config["format"]["colors"][::-1]
starts = config["format"]["starts_with"]

col_list = list(map(hex_to_rgb, col_list))


# Section B: Checking for Updates
status = ""
url = ""

if config["update_check"]:
    try:
        response = requests.get("https://api.github.com/repos/"
            "berkeleyfx/mic-drop-results/releases/latest", timeout=3)

        with open("response.json", "w") as fp:
            json.dump(response.json(), fp)

        version = float(response.json()["tag_name"][1:])
        
        if version > config["version"]:
            print(f"A new version is available. "
                "You can download it using the link below.")
            
            print(f"\nVersion {version}")
            print(response.json()["body"].partition("\n")[0])
            
            url = "https://github.com/berkeleyfx/mic-drop-results/releases/latest/"
            link(url)
            print()

            status = "update available"
        elif version < config["version"]:
            status = "beta"
        else:
            status = "latest version"
        
        status = " [" + status + "]"
    except:
        pass  # Move on if there is no internet connection

print(f"Mic Drop Results (v{config['version']}){status}")

if not "update available" in status:
    url = "https://github.com/berkeleyfx/mic-drop-results"
    link(url)


# Section C: Fixing Command Prompt issues
# Handle KeyboardInterrupt: automatically open the only link
signal.signal(signal.SIGINT, signal.SIG_IGN)

# Disable pausing (QuickEdit and Insert modes)
kernel32 = ctypes.windll.kernel32
kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), 128)


# Section D: Data Cleaning
xls = pd.ExcelFile("data.xlsx")

sheetnames_raw = [n for n in xls.sheet_names if n.lower() != "contestants"]
sheetnames = [re.sub(r'[\\\/:"*?<>|]+', "", name) for name in sheetnames_raw]
data = {}

for i, sheet in enumerate(sheetnames_raw):
    data[sheetnames[i]] = pd.read_excel(xls, sheet)

for k, df in data.items():
    # Check for cases where avg and std are the same (hold the same rank)
    df["r"] = pd.DataFrame(zip(df.iloc[:, 0], df.iloc[:, 1] * -1)) \
        .apply(tuple, axis=1).rank(method="min", ascending=False).astype(int)

    # Sort the slides
    df = df.sort_values(by="r", ascending=True)

    # Remove .0 from whole numbers
    format_number = lambda x: str(int(x)) if x % 1 == 0 else str(x)
    df.loc[:, df.dtypes == float] = df.loc[:, df.dtypes == float].applymap(format_number)

    # Replace {sheet} with sheet name
    df["sheet"] = k

    # Save df to data dictionary
    data[k] = df


# Section E: To PowerPoint
print("\nGenerating slides...")
print("Please do not click on any PowerPoint windows that may show up in the process.\n")

# Kill all PowerPoint instances
subprocess.run("TASKKILL /F /IM powerpnt.exe",
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

# Open template presentation
path = str(pathlib.Path().resolve()) + "\\"
outpath = path + "output\\"
os.makedirs(outpath, exist_ok=True)

for k, df in data.items():
    with alive_bar(8, title=k, title_length=max(map(len, sheetnames)),
        dual_line=True, spinner="classic") as bar:
        # Open template presentation
        bar.text = "Opening template.pptm"
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Presentations.Open(f"{path}template.pptm")
        bar()

        # Import macros
        bar.text = "Importing macros"
        try:
            ppt.VBE.ActiveVBProject.VBComponents.Import(f"{path}Module1.bas")
        except:
            input("\nERROR: Please open PowerPoint, look up Trust Center Settings, "
                "and make sure Trust access to the VBA project object model is checked.")
        bar()

        # Duplicate slides
        bar.text = "Duplicating slides"
        slides_count = ppt.Run("Count")

        for t in df.loc[:, "template"]:
            ppt.Run("Duplicate", t)
        bar()

        # Delete template slides when done
        ppt.Run("DelSlide", *range(1, slides_count + 1))
        bar()

        # Save as output file
        bar.text = "Saving templates"
        output_filename = f"{k}.pptx"

        ppt.Run("SaveAs", f"{outpath}{output_filename}")
        bar()
        ppt.Quit()
        bar()

        # Replace text
        bar.text = "Filling in judging data"
        prs = Presentation(outpath + output_filename)

        for i, slide in enumerate(prs.slides):
            replace_text(slide, df, i)
        bar()

        # Save
        bar.text = f"Saving as {outpath + output_filename}"
        prs.save(outpath + output_filename)
        bar()


# Section F: Launching the File
print(f"\nExported to {outpath}")
print("Press Enter to open the output folder...")

while True:
    try:
        input()
        break
    except EOFError:
        continue

os.startfile(outpath)
