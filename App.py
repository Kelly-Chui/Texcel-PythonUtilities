import tkinter as tk

from Controller import TexcelController as Controller
from Model import TexcelModel as Model
from View import TexcelView as View
import ctypes
import sys

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Texcel")
        self.resizable(False, False)
        model = Model.TexcelModel()

        view = View.TexcelView(self)
        view.grid(padx=20, pady=20)

        controller = Controller.TexcelController(model, view)

        view.set_controller(controller)

if __name__ == '__main__':
    if ctypes.windll.shell32.IsUserAnAdmin():
        app = App()
        app.mainloop()
    else:
        run_as_admin()
