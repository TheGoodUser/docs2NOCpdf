import tkinter as tk
from tkinter import filedialog

def choose_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    try:
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")],
            title="Select an Excel file"
        )
    except Exception as e:
        return None
    return file_path
