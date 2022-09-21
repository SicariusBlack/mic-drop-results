import ctypes
from io import BytesIO
import json
import numpy as np
import os
import pathlib
import re
import requests
import signal
import subprocess
import sys
import traceback
import webbrowser

import pandas as pd

from alive_progress import alive_bar

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.slide import Slide

import win32com
import win32com.client


def throw(*messages, err_type: str = "error"):
    """Throws a handled error with additional guides and details."""
    if len(messages) > 0:
        messages = list(messages)
        messages[0] = f"\n\n{err_type.upper()}: {messages[0]}"
        print(*messages, sep="\n\n")

    if err_type.lower() == "error":
        input("\nPress Enter to exit the program...")
        sys.exit(1)
    else:
        input("\nPress Enter to continue...")


def show_exception_and_exit(exc_type, exc_value, tb):
    traceback.print_exception(exc_type, exc_value, tb)

    # Enable QuickEdit
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x40|0x100))

    throw()


def hex_to_rgb(hex):
    hex = hex.lstrip("#")
    return tuple(int(hex[i:i + 2], 16) for i in (0, 2, 4))


def replace_text(slide: Slide, df, i) -> Slide:
    """Replaces and formats text."""
    cols = df.columns.values.tolist() + ["p"]
    uid = df["uid"].iloc[i]

    for shape in slide.shapes:
        if not shape.has_text_frame or not "{" in shape.text:
            continue

        text_frame = shape.text_frame

        for run in [p.runs[0] for p in text_frame.paragraphs]:
            for search_str in set(re.findall(r"(?<={)(.*?)(?=})", run.text)).intersection(cols):
                if search_str == "p" and uid != np.nan:
                    run.text = ""

                    # Load image from link
                    avatar_url = get_avatar(uid)

                    if avatar_url is None:
                        continue

                    response = requests.get(avatar_url)

                    new_shape = slide.shapes.add_picture(
                        BytesIO(response.content),
                        shape.left, shape.top, shape.width, shape.height
                    )
                    new_shape.auto_shape_type = MSO_SHAPE.OVAL
                    old = shape._element
                    new = new_shape._element
                    old.addnext(new)
                    old.getparent().remove(old)
                    continue

                repl = str(df[search_str].iloc[i])
                repl = repl if repl != "nan" else ""  # Replace missing values with blank

                run.text = run.text.replace("{" + search_str + "}", repl)

                if not search_str.startswith(starts) or not run.font.color.type:
                    continue

                if run.font.color.type == MSO_COLOR_TYPE.RGB:
                    if not run.font.color.rgb == RGBColor(255, 255, 255):
                        continue

                for ind, val in enumerate(range_list):
                    if float(repl) >= val:
                        run.font.color.rgb = RGBColor(*color_list[ind])
                        break
    return slide


def get_avatar(id):
    header = {
        "Authorization": "Bot MTAyMTU5OTE3ODQyMzUzMzY0OQ.Gkx8DG.y_wbRnnf0Nog1UfnpDbGgPellwMi72JyfY5MxU"
    }

    response = requests.get(f"https://discord.com/api/v9/users/{id}", headers=header)

    link = None
    try:
        link = f"https://cdn.discordapp.com/avatars/{id}/{response.json()['avatar']}"
    except KeyError:
        pass

    return link


# Section A: Fixing Command Prompt issues
# Handle KeyboardInterrupt: automatically open the only link
signal.signal(signal.SIGINT, signal.SIG_IGN)

# Disable QuickEdit and Insert mode
kernel32 = ctypes.windll.kernel32
kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x00|0x100))

# Avoid exiting the program when an error is thrown
sys.excepthook = show_exception_and_exit


# Section B: Check if all files are present
missing = [f for f in ["config.json", "data.xlsx", "template.pptm", "Module1.bas"]
    if not os.path.isfile(f)]
    
if len(missing) > 0:
    throw("The following files are missing. Please review the documentation for more "
        "information related to file requirements.", "\n".join(missing))


# Section C: Loading config.json
config = json.load(open("config.json"))

# Variable shortcuts
range_list = config["format"]["ranges"][::-1]
color_list = config["format"]["colors"][::-1]
starts = config["format"]["starts_with"]

color_list = list(map(hex_to_rgb, color_list))


# Section D: Checking for Updates
status = ""
url = ""

if config["update_check"]:
    try:
        response = requests.get("https://api.github.com/repos/"
            "berkeleyfx/mic-drop-results/releases/latest", timeout=3)

        version = float(response.json()["tag_name"][1:])
        
        if version > config["version"]:
            print(f"Version {version}")
            print(response.json()["body"].partition("\n")[0])
            
            url = "https://github.com/berkeleyfx/mic-drop-results/releases/latest/"
            print(url + "\n")
            webbrowser.open(url, new=2)

            status = "update available"
        elif version < config["version"]:
            status = "beta"
        else:
            status = "latest"
        
        status = " [" + status + "]"
    except:
        pass  # Move on if there is no internet connection

print(f"Mic Drop Results (v{config['version']}){status}")

if not "update available" in status:
    url = "https://github.com/berkeleyfx/mic-drop-results"
    print(url)


# Section E: Data Cleaning
path = str(pathlib.Path().resolve()) + "\\"
outpath = path + "output\\"

xls = pd.ExcelFile("data.xlsx")

sheetnames_raw = xls.sheet_names
sheetnames = [re.sub(r'[\\\/:"*?<>|]+', "", name) for name in sheetnames_raw]
data = {}

contestants = None
for s in sheetnames_raw:
    if s.lower() == "contestants":
        contestants = pd.read_excel(xls, s)

for i, sheet in enumerate(sheetnames_raw):
    df = pd.read_excel(xls, sheet)

    # Validate shape
    if df.empty or df.shape < (1, 2):
        continue

    # Exclude contestant database
    if sheet.lower() == "contestants":
        continue

    # Exclude sheets with first two columns where data types are not numeric
    if sum([df.iloc[:, i].dtype.kind in "fuckbitch" for i in range(2)]) < 2:
        throw(f"Invalid data type. The following rows of {sheet} contain string "
            "instead of the supposed numeric data type within the first two columns. "
            "The sheet will be skipped for now.",

            df[~df.iloc[:, :2].applymap(np.isreal).all(1)],

            err_type="warning"
        )

        continue

    # Replace NaN values within the first two columns with 0
    if df.iloc[:, :2].isnull().values.any():
        throw(f"The following rows of {sheet} contain empty values "
            "within the first two columns.",

            df[df.iloc[:, :2].isnull().any(axis=1)],

            "You may exit this program and modify the data or continue with "
            "these values substituted with 0."
            "\nNOTE: Please exit this program before modifying or "
            "Microsoft Excel will throw a sharing violation error.",

            err_type="warning"
        )

        df.iloc[:, :2] = df.iloc[:, :2].fillna(0)

    if contestants is not None:
        df = df.merge(contestants, on="name", how="left")
    
    print(df)

    data[sheetnames[i]] = df

if len(data) < 1:
    throw(f"No valid sheet was found in {path}data.xlsx")

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


# Section F: To PowerPoint
print("\nGenerating slides...")
print("Please do not click on any PowerPoint windows that may show up in the process.\n")

# Kill all PowerPoint instances
subprocess.run("TASKKILL /F /IM powerpnt.exe",
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

# Open template presentation
os.makedirs(outpath, exist_ok=True)

access_error = False
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
            access_error = True
            break     
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

# Warns the user about trust access error
if access_error:
    throw("Please open PowerPoint, look up Trust Center Settings, "
        "and make sure Trust access to the VBA project object model is enabled.")


# Section G: Launching the File
print(f"\nExported to {outpath}")

# Enable QuickEdit
kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x40|0x100))

input("Press Enter to open the output folder...")
os.startfile(outpath)
