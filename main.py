import os
import docx
import pypandoc
import pandas as pd
import win32com.client as win32
from win32com.client import constants
from subprocess import Popen
from utils import *

def doc_to_docx(file_path, del_list):
    def convert_to_docx(file_path_convert, name, del_list):
        temp_name = f'wteTemp{name}'
        # Load the DOCX document
        pypandoc.convert_file(file_path_convert, 'plain', outputfile=f'{temp_name}.txt')
        document = docx.Document()
        
        # Read the text from the file and replace soft returns with paragraph marks
        with open(f'{temp_name}.txt', 'r', encoding='utf-8') as file:
            text = file.readlines()
        for line in text:
            document.add_paragraph(line)
        document.save(f'{temp_name}.docx')
        os.remove(f'{temp_name}.txt')
        del_list.append(os.path.abspath(f'{temp_name}.docx'))
        return del_list
    document = docx.Document(file_path)
    for paragraph in document.paragraphs:
        highlighted_text = extract_format_text(paragraph)
    # Split the file path into name and extension
    name, ext = os.path.splitext(os.path.basename(file_path))
    # Rename path with .docx
    new_file_abs = f"wteTemp{name}.docx"
    if ext == ".doc":
        # Opening MS Word
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(file_path)
        doc.Activate()
        # Save and Close
        word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
        doc.Close(False)
        del_list = convert_to_docx(file_path, name, del_list)
    elif ext == ".docx":
        del_list = convert_to_docx(file_path, name, del_list)
    else: 
        return False, del_list
    return new_file_abs, del_list

def question_create(doc, current_question, current_options, highlights, data, platform, selected_options, question_numbers):
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        # Check if the paragraph is empty
        if not text:
            continue
        
        if is_question(text):
            print(text, 1)
            if current_question and len(current_options) > 0:
                current_question, current_options, highlights = process_options(current_question, current_options, highlights, selected_options, question_numbers)
                question_numbers += 1
                create_quiz(data, current_question, current_options, highlights, platform)
            current_question = text
            current_options = []  # Clear the options list for the new questions
        elif is_option(text):
            highlighted_text = extract_format_text(paragraph)
            highlights.append(highlighted_text)
            for option in split_options(text):
                current_options.append(option)

    question_numbers = last_question(current_question, current_options, highlights, data, platform, selected_options, question_numbers)
    return question_numbers

def last_question(current_question, current_options, highlights, data, platform, selected_options, question_numbers):
    if current_question and current_options:
        question_numbers += 1
        current_question, current_options, highlights = process_options(current_question, current_options, highlights, selected_options, question_numbers)
        create_quiz(data, current_question, current_options, highlights, platform)
    return question_numbers

def data_frame(data, file_path, selected_options):
    df = pd.DataFrame(data)
    if "Xáo trộn câu hỏi" in selected_options:
        df = df.sample(frac=1)  # frac=1 shuffles all rows randomly
    
    # Get the file name without extension
    file_name = os.path.splitext(os.path.basename(file_path))[0] + ".xlsx"
    
    try:
        close_excel(file_path)
        df.to_excel(file_name, index=False)
        os.startfile(file_name)
        Popen(f'explorer /select,"{file_name}"')
    except Exception:
        pass
