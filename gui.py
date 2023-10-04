import docx
import tkinter as tk
from main import open_folder, questionCreate, dataFrame, errorInitiates


def run():
    file_paths = open_folder()  # Returns a tuple of selected file paths
    platform = platform_selection.get()
    selected_options = [option for option, var in checkboxes.items() if var.get()]
    selected_options.extend(s_checkbox_options)
    merge_files = s_var.get()  # Check if the "Gộp nhiều tệp thành một" checkbox is selected

    # Initialize a list to collect data from all selected files
    all_data = []
    question_numbers = 1

    for file_path in file_paths:
        data = []
        current_question = ""
        current_options = []
        highlights = []

        doc = docx.Document(file_path)
        question_numbers = questionCreate(doc, current_question, current_options, highlights, data, platform, selected_options, question_numbers)
        # Append the data to the list if not merging files
        if not merge_files:
            dataFrame(data, file_path)
            error_label.config(text = errorInitiates(code=True))
            
        else:
            all_data.extend(data)  # Collect data from all selected files
            error_label.config(text = errorInitiates(code=True))
            
    # Create a single Excel file containing the combined data if merging files
    if merge_files:
        dataFrame(all_data, "Merged_File.xlsx")
    
    status_label.config(text = "Conversion completed successfully!")
    window.after(2000, window.quit)

# Create the main window
window = tk.Tk()
window.title("Word To Excel Converter")
window.geometry("500x280")  # Increased the height to accommodate the radio buttons

# Main frame for organizing widgets
main_frame = tk.Frame(window)
main_frame.pack(pady=20, padx=10)

# Load the logo image
try:
    p1 = tk.PhotoImage(file='Images\logo.png')
    window.iconphoto(False, p1)
except Exception:
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
checkbox_options = ["Xóa chữ 'Câu'", "Xóa chữ 'A,B,C,D'", "Sửa lỗi định dạng"]
checkboxes = {}

for i, option_text in enumerate(checkbox_options):
    var = tk.BooleanVar()
    checkboxes[option_text] = var
    checkbox = tk.Checkbutton(main_frame, text=option_text, variable=var, anchor="w")
    checkbox.grid(row=3, column=i, pady=10, padx=10, sticky="w")

# Set "Sửa lỗi định dạng" checkbox to be always checked
checkboxes["Sửa lỗi định dạng"].set(True)

# Additional checkbox for merging files
s_checkbox_options = ["Gộp nhiều tệp thành một"]
s_var = tk.BooleanVar()
checkbox = tk.Checkbutton(main_frame, text=s_checkbox_options[0], variable=s_var, anchor="center")
checkbox.grid(row=4, column=0, pady=10, padx=10, sticky="w")

# Status label
status_label = tk.Label(main_frame, text="", fg="green")
status_label.grid(row=5, column=0, columnspan=3, pady=10, padx=10)  # Center the label using "sticky"

# Error label
error_label = tk.Label(main_frame, text="", fg="red")
error_label.grid(row=6, column=0, columnspan=3, pady=10, padx=10)  # Center the label using "sticky"
# Note label in the bottom-right corner with padding
note_label = tk.Label(window, text="caphefalumi v2.1", fg="blue", anchor="w", padx=10)
note_label.pack(side="bottom")

# Start the GUI application
window.mainloop()
