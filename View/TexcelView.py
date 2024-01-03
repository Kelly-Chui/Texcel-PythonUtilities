# text_to_excel_view.py
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from Controller import TexcelController as Controller


class TexcelView(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        # frames
        self.controller = None
        load_path_frame = ttk.Frame(self)
        save_path_frame = ttk.Frame(self)

        # labels
        self.load_path_label = ttk.Label(load_path_frame, text="Select load path", width=40)
        self.save_path_label = ttk.Label(save_path_frame, text="Select save path", width=40)

        # buttons
        load_path_select_button = ttk.Button(load_path_frame, text="..", width=5, command=self.select_load_clicked)
        save_path_select_button = ttk.Button(save_path_frame, text="..", width=5, command=self.select_save_clicked)
        run_button = ttk.Button(self, text="Run", width=10, command=self.run_clicked)
        open_explorer_button = ttk.Button(self, text="open result folder", width=10, command=self.open_explorer_clicked)

        # button styles

        style = ttk.Style()
        style.configure("runButton.TButton", foreground="green", focuscolor="none", padding=(0, 10))
        run_button.config(style="runButton.TButton")

        style.configure("pathLabel.TLabel", background="white", justify="center")
        self.load_path_label.config(style="pathLabel.TLabel")
        self.save_path_label.config(style="pathLabel.TLabel")
        # layouts

        # main Frame
        load_path_frame.grid(row=0, column=0)
        save_path_frame.grid(row=1, column=0)

        run_button.grid(row=0, column=1, rowspan=2)

        self.load_path_label.grid(row=0, column=0, padx=5, pady=5)
        load_path_select_button.grid(row=0, column=1, padx=5, pady=5)
        self.save_path_label.grid(row=0, column=0, padx=5, pady=5)
        save_path_select_button.grid(row=0, column=1, padx=5, pady=5)

        # color settings

    # Commands
    def set_controller(self, controller):
        self.controller = controller

    def select_load_clicked(self):
        self.controller.select_load_clicked()

    def select_save_clicked(self):
        self.controller.select_save_clicked()

    def run_clicked(self):
        self.controller.run_clicked()

    def open_explorer_clicked(self):
        self.controller.open_explorer_clicked()
