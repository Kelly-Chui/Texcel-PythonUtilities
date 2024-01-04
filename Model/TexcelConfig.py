import configparser
import os

class Config:
    def __init__(self):
        try:
            if not os.path.exists("C:/Program Files/Texcel"):
                os.makedirs("C:/Program Files/Texcel")
        except OSError:
            print("Error: Failed to create the directory.")
        self.config_file_path = "C:/Program Files/Texcel/config.ini"
        self.save_path = ""
        self.load_path = ""
        config = configparser.ConfigParser()
        if os.path.exists(self.config_file_path):
            config.read(self.config_file_path)
            self.save_path = config.get("Settings", "save_path", fallback="")
            self.load_path = config.get("Settings", "load_path", fallback="")
        else:
            self.save_path = "Select save path"
            self.load_path = "Select load path"

    def save_last_values(self, save_path, load_path):
        config = configparser.ConfigParser()
        config.read(self.config_file_path)
        if not config.has_section("Settings"):
            config.add_section("Settings")
        config.set("Settings", "save_path", save_path)
        config.set("Settings", "load_path", load_path)

        with open(self.config_file_path, "w") as config_file:
            config.write(config_file)
