import docx
import pandas as pd
import os
import win32com.client
from docx.enum.text import WD_COLOR_INDEX as color
import docx.enum.text



# Open the Word document
doc = docx.Document(r"C:\Users\Toan\WordToExcel\đáp-án-sử-12.docx")
data = []
current_question = ""
options = []
highlights = []


def extract_highlighted_text(paragraph):
    highlighted_text = ""
    for run in paragraph.runs:
        if run.font.highlight_color:  # RGB for yellow
            highlighted_text += run.text
    return highlighted_text

def get_correct_answer_index(options, highlights):
    for i, option_text in enumerate(options):
        if option_text in highlights:
            return i+1
    return None

def close_excel():
    file_path = r"C:\Users\Toan\WordToExcel\questions.xlsx"
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Optional: Hide Excel window
        workbook = excel.Workbooks.Open(file_path)
        workbook.Close(True)  # True to save changes, False to discard changes
        excel.Quit()
        print(f'Closed Excel file: {file_path}')
    except Exception as e:
        print(f'Error: {str(e)}')

# Iterate through the document to extract highlighted text
for paragraph in doc.paragraphs:
    highlighted_text = extract_highlighted_text(paragraph)
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

# Add the last question
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
df = pd.DataFrame(data)
# Save the DataFrame to an Excel file
try:
    close_excel()
    print("Launching....")
    df.to_excel('questions.xlsx', index=False)
    print("Done")
    os.startfile("questions.xlsx")
except Exception as exception:
    print(exception)