import pandas as pd
from os import path
from subprocess import Popen
from utils import *

# Iterate through the document to extract highlighted text and create a quiz
def questionCreate(doc, current_question, current_options, highlights, data, platform):
    for paragraph in doc.paragraphs:
        highlighted_text = extract_format_text(paragraph)
        highlights.append(highlighted_text)

        text = paragraph.text.strip()

        # Check if the paragraph is empty
        if not text:
            continue

        if text.startswith("Câu "):
            # Save the previous question's options and add a new question
            if current_question and current_options:
                while len(current_options) < 4:
                    current_options.append("")  # Fill in empty options if there are less than 4
                create_quiz(data, current_question, current_options, highlights, platform)
            current_question = text
            current_options = []  # Clear the options list for the new question

        # Add the options
        elif is_option(text):
            # Split the options if multiple are on the same line
            for option in split_options(text):
                current_options.append(option)

    # Add the last question if it exists
    lastQuestion(current_question, current_options, highlights, data, platform)
    cls()

# Add the last question and create a quiz
def lastQuestion(current_question, current_options, highlights, data, platform):
    if current_question and current_options:
        while len(current_options) < 4:
            current_options.append("")  # Fill in empty options if there are less than 4
        create_quiz(data, current_question, current_options, highlights, platform)

# Create a DataFrame from the extracted data and save it as an Excel file
def dataFrame(data, file_path):
    df = pd.DataFrame(data)
    
    # Get the file name without extension
    file_name = path.splitext(path.basename(rf'{file_path}'))[0]    
    
    try:
        close_excel()
        df.to_excel(f'{file_name}.xlsx', index=False)
        Popen(rf'explorer /select,"{file_name}.xlsx"')
        os.startfile(f'{file_name}.xlsx')
    except Exception:
        pass
