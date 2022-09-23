import ctypes
from io import BytesIO
import json
import numpy as np
import os
import pathlib
from PIL import Image
import re
import requests
import signal
import subprocess
import sys
import traceback
from urllib.request import Request, urlopen
import webbrowser

import cv2
import pandas as pd

from alive_progress import alive_bar

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.slide import Slide

import win32com
import win32com.client


def is_number(n):
    """Returns True is string is a number."""
    try:
        float(n)
        return True
    except ValueError:
        return False


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

    for shape in slide.shapes:
        if not shape.has_text_frame or not "{" in shape.text:
            continue

        text_frame = shape.text_frame

        for run in [p.runs[0] for p in text_frame.paragraphs]:
            for search_str in set(re.findall(r"(?<={)(.*?)(?=})", run.text)).intersection(cols):
                # Profile picture
                if search_str == "p":
                    effect = run.text[3:].replace(" ", "")
                    if is_number(effect):
                        effect = int(effect)
                    else:
                        effect = 0

                    run.text = ""
                    
                    if db is None or not avatar_mode:
                        continue

                    if df["uid"].iloc[i] == "nan" or not str(df["uid"].iloc[i]).startswith("_"):
                        continue

                    uid = df["uid"].iloc[i][1:]


                    img_path = avapath + str(effect) + "_" + str(uid) + ".png"

                    if not os.path.isfile(img_path):
                        # Load image from link
                        avatar_url = get_avatar(uid)

                        if avatar_url is None:
                            continue

                        avatar_url += ".png"

                        req = urlopen(Request(avatar_url, headers={"User-Agent": "Mozilla/5.0"}))
                        arr = np.asarray(bytearray(req.read()), dtype=np.uint8)
                        img = cv2.imdecode(arr, -1)

                        match effect:
                            case 1:
                                img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                        
                        cv2.imwrite(img_path, img)

                    new_shape = slide.shapes.add_picture(
                        img_path, shape.left, shape.top,
                        shape.width, shape.height
                    )

                    new_shape.auto_shape_type = MSO_SHAPE.OVAL
                    old = shape._element
                    new = new_shape._element
                    old.addnext(new)
                    old.getparent().remove(old)
                    continue

                # Actual text
                repl = str(df[search_str].iloc[i])
                repl = repl if repl != "nan" else ""  # Replace missing values with blank

                run.text = run.text.replace("{" + search_str + "}", repl)

                # Replace image links
                pattern = r"\<\<(.*?)\>\>"
                img_link = re.findall(pattern, run.text)

                if len(img_link) > 0:
                    try:
                        img = BytesIO(requests.get(img_link[0]).content)
                        pil = Image.open(img)

                        im_width = shape.height / pil.height * pil.width
                        new_shape = slide.shapes.add_picture(
                            img, (shape.width - im_width) / 2, shape.top,
                            im_width, shape.height
                        )

                        old = shape._element.addnext(new_shape._element)

                        run.text = ""
                    except:
                        throw("Could not load the following image "
                           f"(Slide {i + 1}, {df['sheet'].iloc[0]}).\n{img_link[0]}",
                            "Please check your internet connection and verify that "
                            "the link leads to an image file. "
                            "It should end with an image extension like .png in most cases.",
                            err_type="warning")

                # Conditional formatting for columns start with "score"
                if not search_str.startswith(starts) or not run.font.color.type:
                    continue

                if run.font.color.type == MSO_COLOR_TYPE.RGB:
                    if not run.font.color.rgb == RGBColor(255, 255, 255):
                        continue

                for ind, val in enumerate(range_list):
                    if is_number(repl):
                        if float(repl) >= val:
                            run.font.color.rgb = RGBColor(*color_list[ind])
                            break
    return slide


def get_avatar(id):
    header = {
        "Authorization": "Bot " + api_token
    }

    if not is_number(id):
        return None

    link = None

    try:
        response = requests.get(f"https://discord.com/api/v9/users/{id}", headers=header)
        link = f"https://cdn.discordapp.com/avatars/{id}/{response.json()['avatar']}"
    except KeyError:
        if response.json()["message"] == "401: Unauthorized":
            throw("Invalid token. Please provide a new token in config.json.",
                response.json())
        else:
            throw(response.json(), err_type="warning")
    except requests.exceptions.ConnectionError:
        global avatar_mode
        avatar_mode = 0
        throw("Please check your internet connection and try again.",
            "Avatars downloading will be skipped for now.", err_type="warning")
    else:
        pass  # Skip all remaining errors

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
avatar_mode = config["avatars"]

with open("token.txt") as f:
    api_token = f.read().splitlines()[0].strip('"')

if len(api_token) < 30 and avatar_mode:
    throw("Please a valid bot token in token.txt or turn off avatar mode in config.json.")

color_list = list(map(hex_to_rgb, color_list))


# Section D: Checking for Updates
status = ""
url = ""

try:
    if config["update_check"]:
        response = requests.get("https://api.github.com/repos/"
            "berkeleyfx/mic-drop-results/releases/latest", timeout=3)

        raw_ver = response.json()["tag_name"][1:]
        version, config_ver = [tuple(map(int, v.split("."))) for v in 
            [raw_ver, config["version"]]
        ]
        
        if version > config_ver:
            print(f"Version {raw_ver}")
            print(response.json()["body"].partition("\n")[0])
            
            url = "https://github.com/berkeleyfx/mic-drop-results/releases/latest/"
            print(url + "\n")
            webbrowser.open(url, new=2)

            status = "update available"
        elif version < config_ver:
            status = "beta"
        else:
            status = "latest"
        
        status = " [" + status + "]"
except requests.exceptions.ConnectionError:
    pass  # Ignore checking for updates without internet connection

print(f"Mic Drop Results (v{config['version']}){status}")

if not "update available" in status:
    url = "https://github.com/berkeleyfx/mic-drop-results"
    print(url)


# Section E: Data Cleaning
path = str(pathlib.Path().resolve()) + "\\"
outpath = path + "output\\"
avapath = path + "avatars\\"

xls = pd.ExcelFile("data.xlsx")

sheetnames_raw = xls.sheet_names
sheetnames = [re.sub(r'[\\\/:"*?<>|]+', "", name) for name in sheetnames_raw]
data = {}

db = None
for s in sheetnames_raw:
    if s.lower() == "contestants":
        db = pd.read_excel(xls, s)

        # Validate shape
        if db.empty or db.shape[0] < 1 or db.shape[1] != 2:
            throw("Contestant database is empty or has invalid shape.\n"
                "Profile pictures will be disabled for now.", err_type="warning")
            db = None
            break

        # Validate name
        if db.columns.values.tolist() != ["name", "uid"]:
            throw("Contestant database does not have valid column names. "
                "The supposed column names are 'name' and 'uid'.\n"
                "Profile pictures will be disabled for now.", err_type="warning")
            db = None

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

    if db is not None:
        df = df.merge(db, on="name", how="left")

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
os.makedirs(avapath, exist_ok=True)

for k, df in data.items():
    with alive_bar(8, title=k, title_length=max(map(len, sheetnames)),
        dual_line=True, spinner="classic", enrich_print=False) as bar:
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
            # Warns the user about trust access error
            throw("Please open PowerPoint, look up Trust Center Settings, "
                "and make sure Trust access to the VBA project object model is enabled.")

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


# Section G: Launching the File
print(f"\nExported to {outpath}")

# Enable QuickEdit
kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x40|0x100))

input("Press Enter to open the output folder...")
os.startfile(outpath)
