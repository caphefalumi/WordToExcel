import os, docx, shutil, win32com.client as win32
import tkinter as tk
from main import open_folder, questionCreate, dataFrame


def doc_to_docx(file_path):
    try: 
        # Split the file path into name and extension
        abs_path = os.path.abspath(file_path)
        name, ext = os.path.splitext(abs_path)
        if ext == ".doc":
            print(True)
            # Create a new file path for the copied file with the .docx extension
            new_file_path = name + ".docx"
            # Copy the original file to the new file path
            shutil.copyfile(abs_path, new_file_path)
            return new_file_path, new_file_path
        elif ext == ".docx":
            """word = win32.client.Dispatch("Word.Application")
            doc = word.Documents.Open(abs_path)

            for content_control in doc.ContentControls:
                if content_control.LockContents:  # Check if it's locked
                    content_control.LockContents = False  # Unlock it
            doc.close(False)
            word.Quit()"""
            return file_path, None
        else: 
            return False, None
    except Exception:
        return False, None

def run():
    file_paths = open_folder()  # Returns a tuple of selected file paths
        
    if not file_paths:
        status_label.config(text="Vui lòng chọn ít nhất một file Word", fg="red")
        return  # Exit the function if no files are selected
    
    platform = platform_selection.get()
    selected_options = [option for option, var in checkboxes.items() if var.get()]

    # Initialize a list to collect data from all selected files
    all_data = []
    doc_list = []
    question_numbers = 1

    for file_path in file_paths:
        data = []
        current_question = ""
        current_options = []
        highlights = []
        path, doc_ext = doc_to_docx(file_path)
        doc = docx.Document(path)
        doc_list.append(doc_ext)
        if doc is False:
            status_label.config(text="Sai định dạng, vui lòng chọn file Word!", fg="red")
            break
        question_numbers = questionCreate(doc, current_question, current_options, highlights, data, platform, selected_options, question_numbers)
        # Append the data to the list if not merging files
        if "Gộp nhiều tệp thành một" not in selected_options:
            dataFrame(data, file_path, selected_options)
        else:
            all_data.extend(data)  # Collect data from all selected files

    # Create a single Excel file containing the combined data if merging files
    if "Gộp nhiều tệp thành một" in selected_options:
        dataFrame(all_data, "Merged_File.xlsx", selected_options)

    if doc is not False:
        for doc_ext in doc_list:
            if doc_ext is not None:
                None #os.remove(doc_ext)
        status_label.config(text="Thành công!", fg="green")
        window.after(2000, window.quit)

# Create the main window
window = tk.Tk()
window.title("Word To Excel Converter v2.3")
window.geometry("480x300")  # Increased the height to accommodate the radio buttons

# Main frame for organizing widgets
main_frame = tk.Frame(window)
main_frame.pack(pady=20, padx=10)

# Load the logo image
try:
    p1 = tk.PhotoImage(file = 'logo.png')
    window.iconphoto(False, p1)
except Exception:
    p1 = tk.PhotoImage(file = 'Images\logo.png')
    window.iconphoto(False, p1)
    pass
# Header label
header_label = tk.Label(main_frame, text="Convert Word to Excel", font=("Helvetica", 16))
header_label.grid(row=0, column=0, columnspan=3, pady=10)  # Center the label using "sticky"

# File selection button
file_button = tk.Button(main_frame, text="Select Word Document", command=run)
file_button.grid(row=1, column=0, columnspan=3, pady=10)  # Center the button using "sticky"

# Platform radio buttons
platform_options = ["Quizizz", "Kahoot", "Blooket"]
platform_selection = tk.StringVar(window)
platform_selection.set(platform_options[0])  # Set the default value

# Create radio buttons
platform_quizizz = tk.Radiobutton(main_frame, text="Quizizz", variable=platform_selection, value="Quizizz")
platform_kahoot = tk.Radiobutton(main_frame, text="Kahoot", variable=platform_selection, value="Kahoot")
platform_blooket = tk.Radiobutton(main_frame, text="Blooket", variable=platform_selection, value="Blooket")

# Place the radio buttons side by side
platform_quizizz.grid(row=2, column=0, pady=10, padx=10, sticky="w")
platform_kahoot.grid(row=2, column=1, pady=10, padx=10, sticky="w")
platform_blooket.grid(row=2, column=2, pady=10, padx=10, sticky="w")

# Choice checkboxes
checkbox_options = ["Xóa chữ 'Câu'", "Xóa chữ 'A,B,C,D'", "Sửa lỗi định dạng","Thêm chữ 'Câu'", "Xáo trộn câu hỏi", "Gộp nhiều tệp thành một"]
checkboxes = {}

for i, option_text in enumerate(checkbox_options):
    var = tk.BooleanVar()
    checkboxes[option_text] = var
    checkbox = tk.Checkbutton(main_frame, text = option_text, variable = var, anchor = "w")
    checkbox.grid(row= 3 + (i // 3), column= i % 3, pady = 10, padx = 10, sticky =" w")

# Set "Sửa lỗi định dạng" checkbox to be always checked
checkboxes["Sửa lỗi định dạng"].set(True)

#Create a frame for the version label
version_label = tk.Label(main_frame, text = "Author: caphefalumi", fg = "blue", font=("Open sans", 8))
version_label.grid(row = 5, column = 2, sticky = "e", padx = 5, pady = 10)

# Status label
status_label = tk.Label(main_frame, text="", fg="green")
status_label.grid(row = 5, column= 0, columnspan = 3, pady = 10, padx = 10)  # Center the label using "sticky"

# Start the GUI application
window.attributes('-topmost', True)
window.mainloop()
