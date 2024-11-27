import customtkinter as ctk
from tkinter import filedialog, Tk, Text
from tkinter.scrolledtext import ScrolledText
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
import sys


class DocxReaderApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Enhanced Docx Reader")
        self.geometry("900x700")

        # Configure system output encoding to handle Unicode
        if sys.stdout.encoding != 'utf-8':
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')

        # File selection button
        self.file_button = ctk.CTkButton(self, text="Open .docx File", command=self.open_file)
        self.file_button.pack(pady=10)

        # Text display area
        self.text_area = ScrolledText(self, wrap="word", width=100, height=40, bg="#f0f0f0", font=("Arial", 12))
        self.text_area.pack(pady=10, padx=10, fill="both", expand=True)
        self.text_area.config(state="disabled")  # Make it read-only initially

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if file_path:
            self.display_docx(file_path)

    def display_docx(self, file_path):
        document = Document(file_path)

        self.text_area.config(state="normal")
        self.text_area.delete("1.0", "end")  # Clear current content

        # Parse paragraphs
        for paragraph in document.paragraphs:
            self.display_paragraph(paragraph)

        self.text_area.config(state="disabled")  # Make it read-only after updating

    def display_paragraph(self, paragraph):
        # Add indentation and line breaks based on paragraph style
        indent = paragraph.paragraph_format.left_indent.pt if paragraph.paragraph_format.left_indent else 0
        indent_spaces = " " * int(indent / 5)  # Approximate spaces for indentation

        for run in paragraph.runs:
            self.render_run(run, indent_spaces)

        # Add a newline at the end of each paragraph
        self.text_area.insert("end", "\n")

    def render_run(self, run, indent_spaces):
        # Extract styling
        font = run.font
        if font.color and isinstance(font.color.rgb, RGBColor):
            r, g, b = font.color.rgb
            text_color = f"#{r:02X}{g:02X}{b:02X}"
        else:
            text_color = "#000000"
        font_family = font.name if font and font.name else "Arial"
        font_size = font.size.pt if font and font.size else 12

        # Handle highlight color properly
        highlight_color = None
        if font and font.highlight_color:
            if isinstance(font.highlight_color, WD_COLOR_INDEX):
                # Map WD_COLOR_INDEX to RGB or other representations
                highlight_color_map = {
                    WD_COLOR_INDEX.AUTO: "#000000",
                    WD_COLOR_INDEX.BLACK: "#000000",
                    WD_COLOR_INDEX.BLUE: "#0000FF",
                    WD_COLOR_INDEX.BRIGHT_GREEN: "#00FF00",
                    WD_COLOR_INDEX.DARK_BLUE: "#00008B",
                    WD_COLOR_INDEX.DARK_RED: "#8B0000",
                    WD_COLOR_INDEX.DARK_YELLOW: "#B5A42E",
                    WD_COLOR_INDEX.GRAY_25: "#D3D3D3",
                    WD_COLOR_INDEX.GRAY_50: "#808080",
                    WD_COLOR_INDEX.GREEN: "#008000",
                    WD_COLOR_INDEX.PINK: "#FFC0CB",
                    WD_COLOR_INDEX.RED: "#FF0000",
                    WD_COLOR_INDEX.TEAL: "#008080",
                    WD_COLOR_INDEX.TURQUOISE: "#40E0D0",
                    WD_COLOR_INDEX.VIOLET: "#EE82EE",
                    WD_COLOR_INDEX.WHITE: "#FFFFFF",
                    WD_COLOR_INDEX.YELLOW: "#FFFF00",
                }
                highlight_color = highlight_color_map.get(font.highlight_color, None)
            elif hasattr(font.highlight_color, "rgb"):
                highlight_color = f"#{font.highlight_color.rgb:06X}"

        bold = font.bold if font else False
        italic = font.italic if font else False
        underline = font.underline if font else False

        # Handle Unicode in debug output safely
        try:
            debug_text = run.text.encode('utf-8').decode('utf-8')
            print(f"Run text: '{debug_text}', Font: {font_family}, Size: {font_size}, "
                  f"Color: {text_color}, Highlight: {highlight_color}, "
                  f"Bold: {bold}, Italic: {italic}, Underline: {underline}",
                  file=sys.stderr)
        except UnicodeEncodeError:
            # If Unicode encoding fails, print a simplified version
            print(f"Run text: <contains special characters>, Font: {font_family}, Size: {font_size}, "
                  f"Color: {text_color}, Highlight: {highlight_color}, "
                  f"Bold: {bold}, Italic: {italic}, Underline: {underline}",
                  file=sys.stderr)

        # Ensure the run's text is rendered with proper styling
        if run.text.strip():  # Skip empty or whitespace-only runs
            self.text_area.insert(
                "end",
                indent_spaces + run.text,
                self.get_style_tag(font_family, font_size, text_color, highlight_color, bold, italic, underline),
            )

    def get_style_tag(self, font_family, font_size, text_color, highlight_color, bold, italic, underline):
        # Generate a unique tag based on the style properties
        tag_name = f"tag_{font_family}_{font_size}_{text_color}_{highlight_color}_{bold}_{italic}_{underline}"

        # Check if the tag already exists
        if tag_name not in self.text_area.tag_names():
            self.text_area.tag_configure(
                tag_name,
                font=(font_family, int(font_size), "bold" if bold else "normal", "italic" if italic else "roman"),
                foreground=text_color,
                background=highlight_color if highlight_color else None,
                underline=underline
            )

        return tag_name


# Run the app
if __name__ == "__main__":
    app = DocxReaderApp()
    app.mainloop()