import os
import docx
import tkinter as tk
from subprocess import Popen, PIPE
from utils import close_excel
from main import open_folder, question_create, data_frame, format_file

def run() -> None:
    """
    Execute the main processing logic for converting Word documents into quiz data.

    This function performs the following steps:
    1. Opens a file dialog to select Word documents for processing.
    2. Validates the file selection and selected options.
    3. Initializes data storage and other variables.
    4. Closes any open Excel files to prevent errors.
    5. Processes each selected file by formatting it, creating quiz questions, and optionally merging multiple files.
    6. Deletes temporary files created during the conversion process.
    7. Opens the 'Output' directory for viewing the generated files.
    8. Updates the status label to indicate the conversion success and closes the application window.
    """    
    # Get selected file paths
    file_paths = open_folder()

    if not file_paths:
        status_label.config(text="Vui lòng chọn ít nhất một file Word", fg="red")
        return

    # Get platform and selected options
    platform = platform_selection.get()
    selected_options = [option for option, var in checkboxes.items() if var.get()]

    # Initialize data collection
    all_data = []
    del_list = []
    question_numbers = 1
    #Close excel first to prevent any error
    close_excel()
    for file_path in file_paths:
        data = []
        current_question = ""
        current_options = []

        # Convert .doc to .docx if needed
        path, highlights, del_list = format_file(file_path, del_list)
        if path is False:
            status_label.config(text="Lỗi định dạng định dạng, vui lòng chọn file Word!", fg="red")
            break
        doc = docx.Document(path)
        question_numbers = question_create(doc, current_question, current_options, highlights, data, platform, selected_options, question_numbers)

        if "Gộp nhiều tệp thành một" not in selected_options:
            question_numbers = 1
            data_frame(data, file_path, selected_options, open=True)
            
        else:
            all_data.extend(data)

    if "Gộp nhiều tệp thành một" in selected_options:
        data_frame(all_data, "Merged_File.xlsx", selected_options, open=True)

    for temp_file in del_list:
        os.remove(temp_file)
    #Open output directory
    process = Popen(['explorer',"Output"], stdout=PIPE, stderr=PIPE)
    stdout, stderr = process.communicate()
    status_label.config(text="Chuyển đổi thành công!", fg="green")
    window.after(2000, window.quit)

# Create the main window
window = tk.Tk()
window.title("Word To Excel Converter v2.3")
window.geometry("480x300")

# Main frame for organizing widgets
main_frame = tk.Frame(window)
main_frame.pack(pady=20, padx=10)

# Load the logo image
try:
    p1 = tk.PhotoImage(file='logo.png')
except Exception:
    p1 = tk.PhotoImage(file='Images\logo.png')
window.iconphoto(False, p1)

# Header label
header_label = tk.Label(main_frame, text="Convert Word to Excel", font=("Helvetica", 16))
header_label.grid(row=0, column=0, columnspan=3, pady=10)

# File selection button
file_button = tk.Button(main_frame, text="Select Word Document", command=run)
file_button.grid(row=1, column=0, columnspan=3, pady=10)

# Platform radio buttons
platform_options = ["Quizizz", "Kahoot", "Blooket"]
platform_selection = tk.StringVar(window)
platform_selection.set(platform_options[0])

# Create radio buttons
platform_quizizz = tk.Radiobutton(main_frame, text="Quizizz", variable=platform_selection, value="Quizizz")
platform_kahoot = tk.Radiobutton(main_frame, text="Kahoot", variable=platform_selection, value="Kahoot")
platform_blooket = tk.Radiobutton(main_frame, text="Blooket", variable=platform_selection, value="Blooket")

# Place the radio buttons side by side
platform_quizizz.grid(row=2, column=0, pady=10, padx=10, sticky="w")
platform_kahoot.grid(row=2, column=1, pady=10, padx=10, sticky="w")
platform_blooket.grid(row=2, column=2, pady=10, padx=10, sticky="w")

# Choice checkboxes
checkbox_options = ["Xóa chữ 'Câu'", "Thêm chữ 'Câu'", "Sửa lỗi định dạng","Xóa chữ 'A,B,C,D'", "Xáo trộn câu hỏi", "Gộp nhiều tệp thành một"]
checkboxes = {}

for i, option_text in enumerate(checkbox_options):
    var = tk.BooleanVar()
    checkboxes[option_text] = var
    checkbox = tk.Checkbutton(main_frame, text=option_text, variable=var, anchor="w")
    checkbox.grid(row=3 + (i // 3), column=i % 3, pady=10, padx=10, sticky="w")

# Set "Sửa lỗi định dạng" checkbox to be always checked
checkboxes["Sửa lỗi định dạng"].set(True)
checkboxes["Thêm chữ 'Câu'"].set(True)

# Create a frame for the version label
version_label = tk.Label(main_frame, text="Author: caphefalumi", fg="blue", font=("Open sans", 8))
version_label.grid(row=5, column=2, sticky="e", padx=5, pady=10)

# Status label
status_label = tk.Label(main_frame, text="", fg="green")
status_label.grid(row=5, column=0, columnspan=3, pady=10, padx=10)  # Center the label using "sticky"

# Start the GUI application
window.mainloop()