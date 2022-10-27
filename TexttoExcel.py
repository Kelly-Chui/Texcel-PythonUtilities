#
#   CombineText.py
#   PythonUtilties
#
#   Created by Kelly Chui on 2022/09/30
#

import os
import re
from openpyxl import Workbook
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def selectLoadFolder():
    global loadPath
    loadPath = filedialog.askdirectory(initialdir = "/", title = "폴더선택")
    if loadPath == "":
        tk.messagebox.showwarning("경고", "폴더선택하시오")
    else:
        res = os.listdir(loadPath)
        print(res)
        if len(res) == 0:
            messagebox.showwarning("경고", "폴더 내 파일 없다")
        else:
            loadPathLabel.config(text=loadPath)
            for file in res:
                print(loadPath + "/" + file)

def selectSaveFolder():
    global savePath
    savePath = filedialog.askdirectory(initialdir = "/", title = "폴더선택")
    savePathLabel.config(text=savePath)
def makeExcelFile(savePath):
    files = os.listdir(loadPath)
    files = list(filter(lambda x: ".txt" in x, files))
    files.sort(key = lambda x: int(''.join(re.findall(r'\d+', x))))

    row = 1
    for fileName in files:
        column = 1
        file = open(loadPath + "/" + fileName)
        for line in file:
            if "=" in line:
                splits = line.split("=")
                sheet.cell(column, row, float(splits[1].rstrip("\n")))
            column += 1
        row += 1
    resultFile.save(savePath + "/" + "result" + ".xlsx")
    print("end")

resultFile = Workbook()
resultFile.remove(resultFile["Sheet"])
sheet = resultFile.create_sheet(index = 1, title = "result")

rootView = tk.Tk()
rootView.title = "TextToExcel"

rootView.geometry("600x400")
rootView.resizable(False, False)

loadPathLabel = tk.Label(rootView, text = "")
savePathLabel = tk.Label(rootView, text = "")

file = tk.Frame(rootView)
file.pack(fill = "x", padx = 5, pady = 5)

loadPath = ""
savePath = ""
textList = []

loadButton = tk.Button(file, text = "Load", width = 12, padx = 5, pady = 5, command = selectLoadFolder)
saveButton = tk.Button(file, text = "Save", width = 12, padx = 5, pady = 5, command = selectSaveFolder)
proceedButton = tk.Button(rootView, text = "Run", width = 12, padx = 15, pady = 15, command = lambda: makeExcelFile(savePath))

loadButton.pack(padx = 5, pady = 5)
loadPathLabel.pack()
saveButton.pack(padx = 10, pady = 5)
savePathLabel.pack()
proceedButton.pack()

rootView.mainloop()

# resultFile = Workbook()
# resultFile.remove(resultFile["Sheet"])
# sheet = resultFile.create_sheet(index = 1, title = "result")
#
# loadDirectory = "/Users/KellyChui/Desktop/results/"
# saveDirectory = "/Users/KellyChui/Desktop/results/"
# saveFileName = "result"
#
# files = os.listdir(loadDirectory)
# files = list(filter(lambda x: ".txt" in x, files))
# files.sort(key = lambda x: int(''.join(re.findall(r'\d+', x))))
#
# row = 1
# for fileName in files:
#     column = 1
#     file = open(loadDirectory + fileName)
#     for line in file:
#         if "=" in line:
#             splits = line.split("=")
#             sheet.cell(column, row, float(splits[1].rstrip("\n")))
#         column += 1
#     row += 1
#
# resultFile.save(saveDirectory + saveFileName + ".xlsx")
