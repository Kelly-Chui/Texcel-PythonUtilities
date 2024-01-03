# text_to_excel_controller.py
import webbrowser
from tkinter import filedialog


class TexcelController:

    def __init__(self, model, view):
        self.model = model
        self.view = view
        self.view.save_path_label.config(text=self.model.save_path)
        self.view.load_path_label.config(text=self.model.load_path)

    def open_url(self, url):
        webbrowser.open(url)

    def select_load_clicked(self):
        folder_path = filedialog.askdirectory(initialdir="/", title="Open text data")
        if folder_path:
            self.model.load_path = folder_path
            self.change_load_label_text(folder_path)

    def select_save_clicked(self):
        folder_path = filedialog.askdirectory(initialdir="/", title="Select a save directory")
        if folder_path:
            self.model.save_path = folder_path
            self.change_save_label_text(folder_path)

    def change_save_label_text(self, text=""):
        self.view.save_path_label.config(text=text)

    def change_load_label_text(self, text=""):
        self.view.load_path_label.config(text=text)

    def run_clicked(self):
        self.model.make_excel_file()
        self.model.set_last_paths(self.model.save_path, self.model.load_path)

    def make_excel_file(self):
        self.model.load_path = self.view.load_path.get()
        self.model.save_path = self.view.save_path.get()
        self.model.make_excel_file()
        self.view.show_information("Complete", "Complete")

    def open_explorer_clicked(self):
        self.open_url("file://" + self.model.save_path.replace("\\", "/"))

    def show_app_info(self):
        self.view.show_app_info()
