import re
from os import path
import win32com.client
from tkinter.filedialog import askopenfilenames


# Helper function to open a window that specifies a file's path
def open_folder():
    # Opens a file dialog to select a file and returns its path.
    filepath = askopenfilenames()
    return filepath

# Helper function to check if a paragraph starts with an option (A, B, C, D)
def is_option(paragraph):
    # Checks if a paragraph starts with an option (A., B., C., D.).
    return paragraph.startswith(("A.", "B.", "C.", "D."))

# Helper function to split options that are on the same line
def split_options(text):
    # Splits options that are on the same line into a list.
    return re.split(r'\s+(?=[A-D]\.)', text)

def extract_format_text(paragraph):
    # Extracts formatted text (highlighted or bold) from a paragraph.
    format_text = ""
    for run in paragraph.runs:
        if run.font.highlight_color or run.bold:
            format_text += run.text
    return format_text

def get_correct_answer_index(options, highlights):
    # Gets the index of the correct answer from options based on highlighted text.
    for i, option_text in enumerate(options):
        if option_text in highlights:
            return i + 1
    return None

def quizizz(data, current_question, current_options, highlights):
    # Creates a Quizizz-style question and adds it to the data list.
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
    # Creates a Kahoot-style question and adds it to the data list.
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

def blooket(data, current_question, current_options, highlights):
    # Creates a Blooket-style question and adds it to the data list.
    data.append({
        'Question Text': current_question,
        'Answer 1': current_options[0],
        'Answer 2': current_options[1],
        'Answer 3': current_options[2],
        'Answer 4': current_options[3],
        'Time limit': 30,
        'Correct Answer': get_correct_answer_index(current_options, highlights),
    })
    return data

def create_quiz(data, current_question, current_options, highlights, platform):
    # Creates a question based on the specified platform and adds it to the data list.
    if platform == "Quizizz":
        quizizz(data, current_question, current_options, highlights)
    elif platform == "Kahoot":
        kahoot(data, current_question, current_options, highlights)
    elif platform == "Blooket":
        blooket(data, current_question, current_options, highlights)

def process_options(current_question, current_options, selected_options):
    # Function to add a period to the end of the option text if it doesn't have one
    def add_period_to_option(text):
        if not text.endswith('.'):
            return (text.strip() + '.')
        return text.strip()
    
    if "Sửa lỗi định dạng" in selected_options:
        # Capitalize "Câu" if it's not already capitalized
        current_question = current_question.replace('câu', 'Câu')
        pattern = r'Câu (\d+)'
        # Add a period after the number following "Câu" if it's missing
        match = re.search(pattern, current_question)
        r_match = re.search(r'^Câu (\d+)\.', current_question)  
        
        if match and not r_match:
            # Add a period after the number
            current_question = re.sub(pattern, lambda m: f'Câu {m.group(1)}.', current_question, 1)
        # Capitalize the text after "Câu X."
        current_question = re.sub(r'Câu (\d+)\.\s*([a-zA-Z])', lambda match: f'Câu {match.group(1)}. {match.group(2).capitalize()}', current_question)
        # Add a period to the end of each option
        current_options = [add_period_to_option(option) for option in current_options]
    if "Remove 'Câu'" in selected_options:
        current_question = re.sub(r'^Câu \d+\.', '', current_question).strip().capitalize()
    if "Remove 'A,B,C,D'" in selected_options:
        current_options = [re.sub(r'[A-D]\.\s*', '', option).strip().capitalize() for option in current_options]
    return current_question, current_options
def close_excel(file_name):
    if path.exists(file_name):
        # Closes an Excel application if it is open.
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Optional: Hide Excel window
            workbook = excel.Workbooks.Open(file_name)
            workbook.Close(True)  # True to save changes, False to discard changes
            excel.Quit()
        except Exception:
            pass
