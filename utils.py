import win32com.client
from tkinter.filedialog import askopenfilename
import re
import os

# Helper function to open a window that specifies a file's path
def cls():
    return os.system('cls')
def open_folder():
    filepath = askopenfilename()
    return filepath

# Helper function to check if a paragraph starts with an option (A, B, C, D)
def is_option(paragraph):
    return paragraph.startswith(("A.", "B.", "C.", "D."))

# Helper function to split options that are on the same line
def split_options(text):
    return re.split(r'\s+(?=[A-D]\.)', text)

def extract_format_text(paragraph):
    format_text = ""
    for run in paragraph.runs:
        if run.font.highlight_color or run.bold:  # Check if the text is highlighted or bold
            format_text += run.text
    return format_text

def get_correct_answer_index(options, highlights):
    for i, option_text in enumerate(options):
        for highlight in highlights:
            if option_text in highlight:
                return i + 1
    return None

def quizizz(data, current_question, current_options, highlights):
    data.append({
        'Question Text': current_question,
        'Question Type': "Multiple Choice",
        'Option 1': current_options[0],
        'Option 2': current_options[1],
        'Option 3': current_options[2],
        'Option 4': current_options[3],
        'Correct Answer': get_correct_answer_index(current_options, highlights),
        'Time in seconds': 30,
    })
    return data

def kahoot(data, current_question, current_options, highlights):
    data.append({
        'Question': current_question,
        'Answer 1': current_options[0],
        'Answer 2': current_options[1],
        'Answer 3': current_options[2],
        'Answer 4': current_options[3],
        'Time limit': 30,
        'Correct Answer': get_correct_answer_index(current_options, highlights),
        
    })
    return data
def close_excel():
    file_path = os.path.abspath(r"questions.xlsx")
    try:
        print("Closing")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Optional: Hide Excel window
        workbook = excel.Workbooks.Open(file_path)
        workbook.Close(True)  # True to save changes, False to discard changes
        excel.Quit()
        cls()
    except Exception as e:
        print(f'Error: {str(e)}')