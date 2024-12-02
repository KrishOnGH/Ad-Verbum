import customtkinter as ctk
from pages.editor import DocxReaderApp
from pages.home import HomePage

class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Ad Verbum")
        self.geometry("900x700")
        self.configure(fg_color="#000000")

        self.home_page = HomePage(self)
        self.home_page.pack(fill="both", expand=True)

    def open_editor(self, file_path=None):
        self.home_page.pack_forget()
        self.editor_page = DocxReaderApp(self, file_path)
        self.editor_page.pack(fill="both", expand=True)

    def back_to_home(self):
        self.editor_page.pack_forget()
        self.home_page.pack(fill="both", expand=True)

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()