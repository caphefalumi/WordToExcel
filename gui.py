import docx
import tkinter as tk
from main import open_folder, questionCreate, dataFrame 

def run():
    file_path = open_folder()
    data = []
    current_question = ""
    current_options = []
    highlights = []

    if file_path:
        doc = docx.Document(file_path)
        platform = platform_selection.get()
        selected_options = [option for option, var in checkboxes.items() if var.get()]
        questionCreate(doc, current_question, current_options, highlights, data, platform, selected_options)
        dataFrame(data, file_path)
        status_label.config(text="Conversion completed successfully!")
        window.after(2000, window.quit)

# Create the main window
window = tk.Tk()
window.title("Word To Excel Converter")
window.geometry("400x250")  # Increased the height to accommodate the radio buttons

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
platform_quizizz.grid(row=2, column=0, pady=10)
platform_kahoot.grid(row=2, column=1, pady=10)
platform_blooket.grid(row=2, column=2, pady=10)

# Choice checkboxes
checkbox_options = ["Remove 'Câu'", "Remove 'A,B,C,D'"]
checkboxes = {}

for i, option_text in enumerate(checkbox_options):
    var = tk.BooleanVar()
    checkboxes[option_text] = var
    checkbox = tk.Checkbutton(main_frame, text=option_text, variable=var, anchor="center")
    checkbox.grid(row=3, column=i, pady=10, sticky="w")

# Status label
status_label = tk.Label(main_frame, text="", fg="green")
status_label.grid(row=4, column=0, columnspan=3, pady=10)  # Center the label using "sticky"

# Note label in the bottom-right corner with padding
note_label = tk.Label(window, text="", fg="blue", anchor="se", padx=10)
note_label.pack(side="bottom", fill="both", expand=True)
note_label.config(text="caphefalumi v2.1")  # Updated text

# Start the GUI application
window.mainloop()
