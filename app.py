import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES

# Create a TkinterDnD window, instead of a regular Tkinter window
class FileDropApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title('Drag and Drop File Example')
        self.geometry('400x200')

        # Create a Label widget to indicate where to drop files
        self.label = tk.Label(self, text="Drag and drop a file here", width=40, height=10, bg='lightgrey')
        self.label.pack(pady=20)

        # Enable the label to accept files
        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        # event.data contains the file path
        file_path = event.data
        self.label.config(text=f"File dropped: {file_path}")

if __name__ == "__main__":
    app = FileDropApp()
    app.mainloop()
