import numpy as np
import pandas as pd
import pathlib

from pptx import Presentation
from pptx_replace import replace_text

import win32com
import win32com.client

# Section A: Data Cleaning
df = pd.read_excel("data.xlsx")
df = df.sort_values(by=list(df.columns[:2]), ascending=[False, True])
df.index = np.arange(0, len(df))

# Check for cases where avg and std are the same (hold the same rank)
df["r"] = pd.DataFrame(zip(df.iloc[:, 0], df.iloc[:, 1] * -1)) \
    .apply(tuple, axis=1).rank(method="min", ascending=False).astype(int)

# Remove .0 from whole numbers
format_number = lambda x: str(int(x)) if x % 1 == 0 else str(x)
df.loc[:, df.dtypes == float] = df.loc[:, df.dtypes == float].applymap(format_number)

print(df)

# Section B: To PowerPoint
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
ppt.Run("SaveAs", f"{path}{output_filename}")
ppt.Quit()

# Section C: Fill in the Blank
prs = Presentation(output_filename)

for i, slide in enumerate(prs.slides):
    for col in df.columns:
        replace_text(slide, "{" + col + "}", str(df[col].iloc[i]))

prs.save(output_filename)
