import pandas as pd
from os import path, startfile
from subprocess import Popen
from utils import *


# Iterate through the document to extract highlighted text and create a quiz
def questionCreate(doc, current_question, current_options, highlights, data, platform, selected_options, question_numbers):
    for paragraph in doc.paragraphs:
        highlighted_text = extract_format_text(paragraph)
        highlights.append(highlighted_text)
        text = paragraph.text.strip()

        # Check if the paragraph is empty
        if not text:
            continue
        if text.startswith("Câu ") or text[0].isdigit() or text[0:1].isdigit() or "Câu" in text:
            # Save the previous question's options and add a new question
            if current_question and len(current_options) > 0:
                current_question, current_options, highlights = process_options(current_question, current_options, highlights, selected_options, question_numbers)
                question_numbers+=1    
                create_quiz(data, current_question, current_options, highlights, platform)
            current_question = text   
            print(current_question)         
            current_options = []  # Clear the options list for the new questions
        # Add the options
        elif is_option(text):
            # Split the options if multiple are on the same line
            for option in split_options(text):
                current_options.append(option)

    # Add the last question if it exists
    question_numbers = lastQuestion(current_question, current_options, highlights, data, platform, selected_options, question_numbers)
    return question_numbers

# Add the last question and create a quiz
def lastQuestion(current_question, current_options, highlights, data, platform, selected_options, question_numbers):
    if current_question and current_options:
        question_numbers+=1
        current_question, current_options, highlights = process_options(current_question, current_options, highlights, selected_options, question_numbers)
        create_quiz(data, current_question, current_options, highlights, platform)
        return question_numbers

# Create a DataFrame from the extracted data and save it as an Excel file
def dataFrame(data, file_path, selected_options):
    df = pd.DataFrame(data)
    if "Xáo trộn câu hỏi" in selected_options:
        df = df.sample(frac=1)  # frac=1 shuffles all rows randomly
    # Get the file name without extension
    file_name = path.splitext(path.basename(rf'{file_path}'))[0] + ".xlsx"    
    try:
        close_excel(rf"{file_path}")
        df.to_excel(file_name, index=False)
        startfile(file_name)
        Popen(rf'explorer /select,"{file_name}"')
    except Exception:
        pass