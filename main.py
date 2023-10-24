import pypandoc
import pandas as pd
from os import path, startfile
from subprocess import Popen
import win32com.client as win32
from win32com.client import constants
from utils import *


def doc_to_docx(file_path):
    def convert(file_path_convert):
        # Load the DOCX document
        pypandoc.convert_file(file_path_convert, 'plain', outputfile='temp.txt')
        document = docx.Document()
        myfile = open(f'wteTemp{name}.txt', 'r', encoding='utf-8').read()
        document.add_paragraph(myfile)
        document.save(f'wteTemp{name}.docx')
        os.remove('wordToExcelConvertTemp.txt')
        return os.path.abspath(f'wteTemp{name}.docx')
    try: 
        # Split the file path into name and extension
        name, ext = os.path.splitext(file_path)
        if ext == ".doc":
            
            # Opening MS Word
            word = win32.gencache.EnsureDispatch('Word.Application')
            doc = word.Documents.Open(file_path)
            doc.Activate ()
            # Rename path with .docx
            new_file_abs = "wordToExcelTemp" + name + ".docx"
            # Save and Close
            word.ActiveDocument.SaveAs(
                new_file_abs, FileFormat=constants.wdFormatXMLDocument
            )
            doc.Close(False)
            new_file_abs = convert(new_file_abs, name)
            
            return new_file_abs, new_file_abs
        elif ext == ".docx":
            return file_path, None
        else: 
            return False, None
    except Exception:
        return False, None

# Iterate through the document to extract highlighted text and create a quiz
def questionCreate(doc, current_question, current_options, highlights, data, platform, selected_options, question_numbers):
    for paragraph in doc.paragraphs:
        highlighted_text = extract_format_text(paragraph)
        highlights.append(highlighted_text)
        text = paragraph.text.strip()
        print(text)
        # Check if the paragraph is empty
        if not text:
            continue
        if is_question(text):
            # Save the previous question's options and add a new question
            if current_question and len(current_options) > 0:
                current_question, current_options, highlights = process_options(current_question, current_options, highlights, selected_options, question_numbers)
                question_numbers+=1    
                create_quiz(data, current_question, current_options, highlights, platform)
            current_question = text   
            print(current_question)         
            current_options = []  # Clear the options list for the new questions
        elif is_option(text):
            # Split the options if multiple are on the same line
            for option in split_options(text):
                current_options.append(option)

        # Add the options


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