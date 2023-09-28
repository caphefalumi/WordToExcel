import pandas as pd
import os
from init import *
import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT



data = []
current_question = ""
current_options = []
highlights = []

# Iterate through the document to extract highlighted text
def questionCreate(doc, current_question, current_options, highlights, data):

    print('Creating...')
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
            current_question = text
            current_options = []  # Clear the options list for the new question
        elif is_option(text):
            # Split the options if multiple are on the same line
            for option in split_options(text):
                current_options.append(option)

    # Add the last question if it exists
    lastQuestion(current_question,current_options,highlights,data)

    cls()

# Add the last question
def lastQuestion(current_question, current_options, highlights, data):
    if current_question and current_options:
        while len(current_options) < 4:
            current_options.append("")  # Fill in empty options if there are less than 4
        
        data.append({
            'Question Text': current_question,
            'Question Type': "Multiple Choice",
            'Option 1': current_options[0],
            'Option 2': current_options[1],
            'Option 3': current_options[2],
            'Option 4': current_options[3],
            'Correct Answer': get_correct_answer_index(current_options, highlights),  # No correct answer specified in the provided format
            'Time in seconds': 30,  # You can set the time as needed
        })

#dùng def bằng cách split
# Create a DataFrame from the extracted data
def dataFrame(data):
    df = pd.DataFrame(data)
    # Save the DataFrame to an Excel file
    try:
        close_excel()
        df.to_excel('questions.xlsx', index=False)
        os.startfile("questions.xlsx")
    except Exception as exception:
        print(exception)