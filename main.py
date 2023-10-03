import pandas as pd
from os import path, startfile
from subprocess import Popen
from utils import *

# Iterate through the document to extract highlighted text and create a quiz
def questionCreate(doc, current_question, current_options, highlights, data, platform, selected_options, question_number):
    for paragraph in doc.paragraphs:
        highlighted_text = extract_format_text(paragraph)
        highlights.append(highlighted_text)
        text = paragraph.text.strip()

        # Check if the paragraph is empty
        if not text:
            continue

        if text.startswith("Câu ") or text[0].isdigit() or text[0:1].isdigit():
            # Save the previous question's options and add a new question
            if current_question and current_options:
                current_question, current_options = process_options(current_question, current_options, selected_options, question_number)
                question_number+=1                
                create_quiz(data, current_question, current_options, highlights, platform)
            current_question = text
            current_options = []  # Clear the options list for the new question

        # Add the options
        elif is_option(text):
            # Split the options if multiple are on the same line
            for option in split_options(text):
                current_options.append(option)

    # Add the last question if it exists
    question_number = lastQuestion(current_question, current_options, highlights, data, platform, selected_options, question_number)
    return question_number
# Add the last question and create a quiz
def lastQuestion(current_question, current_options, highlights, data, platform, selected_options, question_number):
    if current_question and current_options:
        current_question, current_options = process_options(current_question, current_options, selected_options, question_number)
        create_quiz(data, current_question, current_options, highlights, platform)
    return question_number


# Create a DataFrame from the extracted data and save it as an Excel file
def dataFrame(data, file_path):
    df = pd.DataFrame(data)
    # Get the file name without extension
    file_name = path.splitext(path.basename(rf'{file_path}'))[0] + ".xlsx"    
    try:
        close_excel(file_name)
        df.to_excel(file_name, index=False)
        Popen(rf'explorer /select,"{file_name}"')
        startfile(file_name)
    except Exception:
        pass
