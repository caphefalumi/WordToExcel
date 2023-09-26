import tkinter as tk
import docx
from tkinter import filedialog
from main import *
from init import *



os.system('cls')
print("Launching...")
def Run():
    doc = docx.Document(rf"{open_folder()}")
    format_paragraph(doc)
    temp_doc = docx.Document(r'temp.docx')
    questionCreate(temp_doc,current_question,options,highlights,data)
    dataFrame(data)
    os.remove(r'temp.docx')


window = tk.Tk()
window.title("Word To Excel")

organize_button = tk.Button(window, text="WordToExcel", command=Run)
organize_button.pack()

status_label = tk.Label(window, text="")
status_label.pack()

window.mainloop()