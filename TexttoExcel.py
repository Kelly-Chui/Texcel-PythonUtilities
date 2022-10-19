#
#   CombineText.py
#   PythonUtilties
#
#   Created by Kelly Chui on 2022/09/30
#

import os
import re
from openpyxl import Workbook

resultFile = Workbook()
resultFile.remove(resultFile["Sheet"])
sheet = resultFile.create_sheet(index = 1, title = "result")

loadDirectory = "/Users/KellyChui/Desktop/results/"
saveDirectory = "/Users/KellyChui/Desktop/results/"
saveFileName = "result"

files = os.listdir(loadDirectory)
files = list(filter(lambda x: ".txt" in x, files))
files.sort(key = lambda x: int(''.join(re.findall(r'\d+', x))))

row = 1
for filename in files:
    column = 1
    file = open(loadDirectory + filename)
    for line in file:
        if "=" in line:
            splits = line.split("=")
            sheet.cell(column, row, float(splits[1].rstrip("\n")))
        column += 1
    row += 1

resultFile.save(saveDirectory + saveFileName + ".xlsx")
