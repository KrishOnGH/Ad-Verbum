import customtkinter as ctk
from tkinter import filedialog
import json
import os

class HomePage(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.master = master

        self.open_button = ctk.CTkButton(
            self, 
            text="Open .docx File",
            command=self.open_file,
            fg_color="#333333",
            hover_color="#444444",
            text_color="white"
        )
        self.open_button.pack(pady=10)

        self.create_button = ctk.CTkButton(
            self, 
            text="Create New Document",
            command=self.create_file,
            fg_color="#333333",
            hover_color="#444444",
            text_color="white"
        )
        self.create_button.pack(pady=10)

        self.recent_files_list = ctk.CTkLabel(self, text="Recently Opened Files:")
        self.recent_files_list.pack(pady=10)

        self.recent_files = self.load_recent_files()
        for file in self.recent_files:
            file_button = ctk.CTkButton(
                self, 
                text=file,
                command=lambda f=file: self.master.open_editor(f),
                fg_color="#333333",
                hover_color="#444444",
                text_color="white"
            )
            file_button.pack(pady=5)

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if file_path:
            self.add_to_recent_files(file_path)
            self.master.open_editor(file_path)

    def create_file(self):
        # Logic to create a new document
        pass

    def load_recent_files(self):
        if not os.path.exists("data"):
            os.makedirs("data")
        if os.path.exists("data/recent_files.json"):
            with open("data/recent_files.json", "r") as file:
                return json.load(file)
        return []

    def add_to_recent_files(self, file_path):
        if file_path not in self.recent_files:
            self.recent_files.append(file_path)
            with open("data/recent_files.json", "w") as file:
                json.dump(self.recent_files, file)