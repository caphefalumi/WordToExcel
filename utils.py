import re
import subprocess
from tkinter.filedialog import askopenfilenames

# Helper function to open a window that specifies a file's path
def open_folder():
    # Opens a file dialog to select a file and returns its path.
    filepaths = askopenfilenames()
    return filepaths

# Capitalize first letter
def CFL(text: str) -> str:
    if text:
        return text[0].upper() + text[1:]
    else:
        return text

# Helper function to check whether a text is a question
def is_question(text: str) -> bool:
    if text.startswith("Câu ") or text.startswith("Câu") or re.match(r"(\d+)\.", text):
        return True

# Helper function to check if a paragraph starts with an option (A, B, C, D)
def is_option(text) -> bool:
    if text.startswith(("A.", "B.", "C.", "D.", "a.", "b.", "c.", "d.")):
        return True

# Helper function to split options that are on the same line
def split_options(text: str) -> list:
    # Splits options that are on the same line into a list.
    if is_option(text):
        return re.split(r'\s+(?=[a-dA-D]\.)', text)

def extract_format_text(text) -> str:
    # Extracts formatted text (highlighted or bold) from a paragraph.
    format_text = ""
    for run in text.runs:
        if run.font.highlight_color or run.bold or run.underline or run.italic:
            format_text += run.text
    return format_text

def get_correct_answer_index(options: list, highlights: list) -> int:
    # Gets the index of the correct answer from options based on highlighted text.
    for i, option_text in enumerate(options):
        cleaned_text = re.sub(r'^[a-dA-D]\. ', '', option_text).strip()
        if cleaned_text == highlights[0]:
            highlights.pop(0)
            return i+1
    return None

def create_quiz(data: list, current_question: str, current_options: list, highlights: list, platform: str):
    # Creates a question based on the specified platform and adds it to the data list.
    def quizizz(data: list, current_question: str, current_options: list, highlights: list) -> list:
        # Creates a Quizizz-style question and adds it to the data list.
        data.append({
            'Question Text': current_question,
            'Question Type': "Multiple Choice",
            'Option 1': current_options[0] if len(current_options) > 0 else "",
            'Option 2': current_options[1] if len(current_options) > 1 else "",
            'Option 3': current_options[2] if len(current_options) > 2 else "",
            'Option 4': current_options[3] if len(current_options) > 3 else "",
            'Correct Answer': get_correct_answer_index(current_options, highlights),
            'Time in seconds': 30,
        })
        return data

    def kahoot(data: list, current_question: str, current_options: list, highlights: list) -> list:
        # Creates a Kahoot-style question and adds it to the data list.
        data.append({
            'Question': current_question,
            'Answer 1': current_options[0] if len(current_options) > 0 else "",
            'Answer 2': current_options[1] if len(current_options) > 1 else "",
            'Answer 3': current_options[2] if len(current_options) > 2 else "",
            'Answer 4': current_options[3] if len(current_options) > 3 else "",
            'Time limit': 30,
            'Correct Answer': get_correct_answer_index(current_options, highlights),
        })
        return data

    def blooket(data: list, current_question: str, current_options: list, highlights: list) -> list:
        # Creates a Blooket-style question and adds it to the data list.
        answers = {}
        for i in range(len(current_options)):
            answers[f'Answer {i + 1}'] = current_options[i]

        data.append({
            'Question Text': current_question,
            'Answer 1': current_options[0] if len(current_options) > 0 else "",
            'Answer 2': current_options[1] if len(current_options) > 1 else "",
            'Answer 3': current_options[2] if len(current_options) > 2 else "",
            'Answer 4': current_options[3] if len(current_options) > 3 else "",
            'Time limit': 30,
            'Correct Answer': get_correct_answer_index(current_options, highlights),
        })
        return data

    try:
        if platform == "Quizizz":
            quizizz(data, current_question, current_options, highlights)
        elif platform == "Kahoot":
            kahoot(data, current_question, current_options, highlights)
        elif platform == "Blooket":
            blooket(data, current_question, current_options, highlights)
    except Exception:
        pass

def process_options(current_question: str, current_options: list, selected_options:list, question_number: int):
    pattern = r'Câu (\d+)'
    match = re.search(pattern, current_question)
    r_match_1 = re.search(r'^Câu (\d+)\.', current_question)
    r_match_2 = re.search(r'^Câu (\d+)\:', current_question)
    r_match_3 = re.search(r'^Câu (\d+) ', current_question)
    current_question = current_question.replace('câu', 'Câu')

    if "Sửa lỗi định dạng" in selected_options:
        # Add a period after the number following "Câu" if it's missing
        if match and not r_match_1 and not r_match_2:
            # Add a period after the number
            current_question = re.sub(pattern, lambda m: f'Câu {m.group(1)}.', current_question, 1)

        # Capitalize the text after "Câu X."
        current_question = re.sub(r'Câu (\d+)\.\s*([a-zA-Z])', lambda match: f'Câu {match.group(1)}. {CFL(match.group(2))}', current_question)
        current_options = [re.sub(r'([a-dA-D])\.\s*(.*)', lambda match: f'{CFL(match.group(1))}. {CFL(match.group(2).strip())}', option) for option in current_options]

    if "Xóa chữ 'Câu'" in selected_options:
        current_question = CFL(re.sub(r'^Câu \d+\.', '', current_question).strip())
        current_question = CFL(re.sub(r'^Câu \d+\:', '', current_question).strip())
        current_question = CFL(re.sub(r'\d+\.', '', current_question).strip())

    if "Xóa chữ 'A,B,C,D'" in selected_options:
        current_options = [CFL(re.sub(r'[a-dA-D]\.\s*', '', option).strip()) for option in current_options]

    if "Thêm chữ 'Câu'" in selected_options and not "Câu" in current_question:
        current_question = re.sub(r"(\d+)", r'Câu \1', current_question, 1)

    if match:
        current_question = re.sub(pattern, f"Câu {question_number}", current_question)

    return current_question, current_options

def close_excel():
    # Closes an Excel application if it is open.
    subprocess.call("TASKKILL /F /IM EXCEL.EXE > nul 2>&1", shell=True)
