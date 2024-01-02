import os
import tkinter as tk
import webbrowser
from tkinter import filedialog
from tkinter import messagebox

import natsort
import pandas as pd


def selectLoadFolder():
    global load_path
    load_path = filedialog.askdirectory(initialdir="/", title="Open text data")
    if load_path == "":
        tk.messagebox.showwarning("Warning", "It's not a folder.")
    else:
        res = os.listdir(load_path)
        print(res)
        if len(res) == 0:
            messagebox.showwarning("Warning", "There are no files inside the folder.")
        else:
            loadPathLabel.config(text=load_path)
            for file in res:
                print(load_path + "/" + file)


def selectSaveFolder():
    global savePath
    savePath = filedialog.askdirectory(initialdir="/", title="Select a save directory")
    savePathLabel.config(text=savePath)

def extract_number(x):
    if 'vout=' in x:
        try:
            return float(x.split('=')[1])
        except ValueError:
            return "error"
    else:
        return "error"

def makeExcelFile(save_path):

    if load_path == "" or save_path == "":
        messagebox.showwarning("Notice", "Directories must be selected")
        return

    files = natsort.natsorted([file for file in os.listdir(load_path) if file.endswith(".txt")])
    print(files)
    df_list = []
    row = 1
    for file_idx in range(len(files)):
        with open(os.path.join(load_path, files[file_idx]), 'r') as txt_file:
            file_df = pd.read_csv(txt_file, header=None, names=None)
            df_list.append(file_df)
    result_df = pd.concat(df_list, axis=1)
    result_df = result_df.map(extract_number)

    with pd.ExcelWriter(save_path + "/" + "output.xlsx", engine='openpyxl') as writer:
        result_df.to_excel(writer, index=False, header=False)


def openExplorer():
    os.system(f"explorer {os.path.realpath(savePath)}")


def openURL():
    webbrowser.open("https://github.com/Kelly-Chui/PythonUtilities")


def openpyxlURL():
    webbrowser.open("https://openpyxl.readthedocs.io/en/stable/")


def openNatsortURL():
    webbrowser.open("https://github.com/SethMMorton/natsort")


def showAppinfo():
    appInfoView = tk.Toplevel()
    appInfoView.title("App infomation")
    appInfoView.geometry("200x200")
    appInfoView.resizable(False, False)
    maker = tk.Label(appInfoView, text="Made by Kelly Chui")
    githubURL = tk.Button(appInfoView, text="contact(Website)", command=openURL)
    openSourceLabel = tk.Label(appInfoView, text="Opensource License")
    openSourceInfo1 = tk.Button(appInfoView, text="natsort", comman=openNatsortURL)
    openSourceInfo2 = tk.Button(appInfoView, text="openpyxl", command=openpyxlURL)

    maker.pack(side="top")
    githubURL.pack(side="top")

    openSourceInfo1.pack(side="bottom")
    openSourceInfo2.pack(side="bottom")
    openSourceLabel.pack(side="bottom")


rootView = tk.Tk()
rootView.configure(bg="white")
rootView.title("TextToExcel")
rootView.geometry("600x400")
rootView.resizable(False, False)

file = tk.Frame(rootView)
file.pack(fill="x", padx=5, pady=5)

load_path = ""
savePath = ""
textList = []

loadLabel = tk.Label(rootView, bg="white", height=10)
loadLabel.pack(side="top", pady=10)
saveLabel = tk.Label(rootView, bg="white", height=10)
saveLabel.pack(side="top", pady=10)
runLabel = tk.Label(rootView, bg="white")
runLabel.pack()
infoLabel = tk.Label(rootView, bg="white", height=20)
infoLabel.pack(side="bottom")

loadPathLabel = tk.Label(loadLabel, text="loadPath", width="30")
savePathLabel = tk.Label(saveLabel, text="savePath", width="30")
loadButton = tk.Button(loadLabel, text="Load", bg="white", width=5, command=selectLoadFolder)
saveButton = tk.Button(saveLabel, text="Save", width=5, bg="white", command=selectSaveFolder)
proceedButton = tk.Button(runLabel, text="Run", fg="white", bg="green", width=10, pady=5,
                          command=lambda: makeExcelFile(savePath))
openSaveButton = tk.Button(runLabel, text="Show results in Explorer", bg="white", fg="blue", width=20,
                           command=openExplorer)
infoButton = tk.Button(infoLabel, text="App info", bg="white", fg="green", command=showAppinfo)

loadButton.pack(side="right", padx=10)
loadPathLabel.pack(side="right")
saveButton.pack(side="right", padx=10)
savePathLabel.pack(side="right")
proceedButton.pack(side="top")
openSaveButton.pack(side="top")
infoButton.pack()
rootView.mainloop()
