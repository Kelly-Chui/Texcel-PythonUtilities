import os
from datetime import datetime
from tkinter import messagebox

import natsort
import pandas as pd

from Model import TexcelConfig


class TexcelModel:
    def __init__(self):
        self.config = TexcelConfig.Config()
        self.save_path = self.config.save_path
        self.load_path = self.config.load_path

    def set_last_paths(self, save_path, load_path):
        self.config.save_last_values(save_path, load_path)

    def make_excel_file(self):
        def extract_number(x):
            if "=" in x:
                try:
                    return float(x.split('=')[1])
                except ValueError:
                    return "failed"
            else:
                return "failed"

        if self.load_path == "" or self.save_path == "":
            messagebox.showwarning("Notice", "No valid directory has been selected. Please reselect")
            return

        files = natsort.natsorted([file for file in os.listdir(self.load_path) if file.endswith(".txt")])
        df_list = []
        for file_idx in range(len(files)):
            with open(os.path.join(self.load_path, files[file_idx]), 'r') as txt_file:
                file_df = pd.read_csv(txt_file, header=None, names=None)
                df_list.append(file_df)
        result_df = pd.concat(df_list, axis=1)
        result_df = result_df.map(extract_number)

        now = datetime.now().strftime("%m_%d_%Y_%H_%M_%S")
        with pd.ExcelWriter(os.path.join(self.save_path, f"result_{now}.xlsx"), engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False, header=False)
        messagebox.showinfo("info", "Complete")
