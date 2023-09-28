import tkinter as tk
from tkinter import filedialog
import docx
from main import *
from init import *

def run():
    file_path = open_folder()
    if file_path:
        doc = docx.Document(file_path)
        format_paragraph(doc)
        temp_doc = docx.Document('temp.docx')
        questionCreate(temp_doc, current_question, current_options, highlights, data)
        dataFrame(data)
        #os.remove('temp.docx')
        status_label.config(text="Conversion completed successfully!")
        os.remove(r'temp.docx')


# Create the main window
window = tk.Tk()
window.title("Word To Excel Converter")
window.geometry("400x200")

# Main frame for organizing widgets
main_frame = tk.Frame(window)
main_frame.pack(pady=20)

# Header label
header_label = tk.Label(main_frame, text="Convert Word to Excel", font=("Helvetica", 16))
header_label.grid(row=0, column=0, columnspan=2, pady=10)

# File selection button
file_button = tk.Button(main_frame, text="Select Word Document", command=run)
file_button.grid(row=1, column=0, columnspan=2, pady=10)

# Status label
status_label = tk.Label(main_frame, text="", fg="green")
status_label.grid(row=2, column=0, columnspan=2, pady=10)

# Additional features and widgets can be added here

# Start the GUI application
window.mainloop()