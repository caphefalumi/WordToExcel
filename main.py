import docx
import pandas as pd
import os
from docx.enum.text import WD_COLOR_INDEX as color
import docx.enum.text
from init import *
from time import sleep 

# Open the Word document
data = []
current_question = ""
options = []
highlights = []

# Iterate through the document to extract highlighted text
def questionCreate(doc,current_question, options, highlights, data):
    print('Creating...')
    for paragraph in doc.paragraphs:
        highlighted_text = extract_format_text(paragraph)
        highlights.append(highlighted_text)

        text = paragraph.text.strip()
        # Check if the paragraph is empty
        if not text:
            continue
        # Check if the paragraph starts with "Câu X." to identify a new question

        if text.startswith("Câu "):
            # Save the previous question if there was one
            if current_question and options:
                while len(options) < 4:
                    options.append("")  # Fill in empty options if there are less than 4
                
                data.append({
                    'Question Text': current_question,
                    'Question Type': "Multiple Choice",
                    'Option 1': options[0],
                    'Option 2': options[1],
                    'Option 3': options[2],
                    'Option 4': options[3],
                    'Correct Answer': get_correct_answer_index(options, highlights),  # No correct answer specified in the provided format
                    'Time in seconds': 30,  # You can set the time as needed
                })

            # Start a new question
            current_question = text
            options = []  # Clear the options list for the new question

        else:
            # Append options to the list
            options.append(text)
    lastQuestion(current_question, options, highlights, data)
    cls()
# Add the last question
def lastQuestion(current_question, options, highlights, data):
    if current_question and options:
        while len(options) < 4:
            options.append("")  # Fill in empty options if there are less than 4
        
        data.append({
            'Question Text': current_question,
            'Question Type': "Multiple Choice",
            'Option 1': options[0],
            'Option 2': options[1],
            'Option 3': options[2],
            'Option 4': options[3],
            'Correct Answer': get_correct_answer_index(options, highlights),  # No correct answer specified in the provided format
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