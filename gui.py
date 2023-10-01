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
        questionCreate(doc, current_question, current_options, highlights, data, platform)
        dataFrame(data, file_path)
        status_label.config(text="Conversion completed successfully!")
        window.after(2000, window.quit)

# Create the main window
window = tk.Tk()
window.title("Word To Excel Converter")
window.geometry("400x250")  # Increased the height to accommodate the radio buttons

# Main frame for organizing widgets
main_frame = tk.Frame(window)
main_frame.pack(pady=20)

# Load the logo image
p1 = tk.PhotoImage(file = 'Images\logo.png')
window.iconphoto(False,p1)

# Header label
header_label = tk.Label(main_frame, text="Convert Word to Excel", font=("Helvetica", 16))
header_label.grid(row=0, column=0, columnspan=3, pady=10, sticky="n")  # Center the label using "sticky"

# File selection button
file_button = tk.Button(main_frame, text="Select Word Document", command=run)
file_button.grid(row=1, column=0, columnspan=3, pady=10, sticky="n")  # Center the button using "sticky"

# Platform radio buttons
platform_options = ["Quizizz", "Kahoot", "Blooket"]
platform_selection = tk.StringVar(window)
platform_selection.set(platform_options[0])  # Set the default value

# Create radio buttons
platform_quizizz = tk.Radiobutton(main_frame, text="Quizizz", variable=platform_selection, value="Quizizz")
platform_kahoot = tk.Radiobutton(main_frame, text="Kahoot", variable=platform_selection, value="Kahoot")
platform_blooket = tk.Radiobutton(main_frame, text="Blooket", variable=platform_selection, value="Blooket")

# Place the radio buttons in the center below the Select Word Document button
platform_quizizz.grid(row=2, column=0, pady=10, sticky="n")
platform_kahoot.grid(row=2, column=1, pady=10, sticky="n")
platform_blooket.grid(row=2, column=2, pady=10, sticky="n")

# Status label
status_label = tk.Label(main_frame, text="", fg="green")
status_label.grid(row=3, column=0, columnspan=3, pady=10, sticky="n")  # Center the label using "sticky"

# Note label in the bottom-right corner
note_label = tk.Label(window, text="", fg="blue", anchor="se")
note_label.pack(side="bottom", fill="both", expand=True)
note_label.config(text="caphefalumi\nv2.1") 

# Start the GUI application
window.mainloop()   
