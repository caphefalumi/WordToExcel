import os
import docx
import pypandoc
import pandas as pd
import win32com.client as win32
from win32com.client import constants
from utils import *

def doc_to_docx(file_path: str, del_list: list):
    def convert_to_docx(file_path_convert: str, name: str, del_list: list) -> list:
        temp_name = f'wteTemp{name}'
        # Load the DOCX document
        pypandoc.convert_file(file_path_convert, 'plain', extra_args=['--wrap=none'], outputfile=f'{temp_name}.txt')
        document = docx.Document()
        # Read the text from the file and replace soft returns with paragraph marks
        with open(f'{temp_name}.txt', 'r', encoding='utf-8') as file:
            text = file.readlines()
        for line in text:
            document.add_paragraph(line)
        document.save(f'{temp_name}.docx')
        del_list.append(os.path.abspath(f'{temp_name}.docx'))
        os.remove(f'{temp_name}.txt')
        return del_list
    #Return a list of format text
    def extract_original_format(file_path: str) -> list:
        highlights = []
        document = docx.Document(file_path)
        for paragraph in document.paragraphs:
            highlighted_text = extract_format_text(paragraph)
            if is_option(highlighted_text): 
                highlights.append(CFL(re.sub(r'^[a-dA-D]\.', '', highlighted_text).strip()))
        return highlights
    # Split the file path into name and extension
    name, ext = os.path.splitext(os.path.basename(file_path))
    # Rename path with wteTemp(file name in here).docx
    new_file_abs = f"wteTemp{name}.docx"
    if ext == ".doc":
        temp_name = f"wteDocTemp{name}"
        temp_path = os.path.abspath(f"{temp_name}.docx")
        # Opening MS Word
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(os.path.abspath(file_path))
        doc.Activate()
        # Save and Close
        word.ActiveDocument.SaveAs(temp_name, FileFormat=constants.wdFormatXMLDocument)
        doc.Close(True)
        del_list = convert_to_docx(temp_path, name, del_list)
        del_list.append(temp_path)
        highlights = extract_original_format(temp_path)
        return new_file_abs, highlights, del_list
    #If it is .docx dont need to convert to .docx again
    elif ext == ".docx":
        del_list = convert_to_docx(file_path, name, del_list)
        highlights = extract_original_format(file_path)
        return new_file_abs, highlights, del_list   
    return False, None, None

def question_create(doc, current_question: str, current_options: list, highlights: list, data: list, platform: str, selected_options: list, question_numbers: int) -> int: #Return question numbers
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        # Check if the paragraph is empty
        if not text:
            continue
        if is_question(text):
            if current_question and len(current_options) > 0:
                current_question, current_options = process_options(current_question, current_options, selected_options, question_numbers)
                question_numbers += 1
                create_quiz(data, current_question, current_options, highlights, platform)
            current_question = text
            current_options = []  # Clear the options list for the new questions
        elif is_option(text):
            for option in split_options(text):
                current_options.append(option)

    question_numbers = last_question(current_question, current_options, highlights, data, platform, selected_options, question_numbers)
    return question_numbers

def last_question(current_question: str, current_options: list, highlights: list, data: list, platform: str, selected_options: list, question_numbers: int) -> int: #Return question numbers
    if current_question and current_options:
        question_numbers += 1
        current_question, current_options = process_options(current_question, current_options, selected_options, question_numbers)
        create_quiz(data, current_question, current_options, highlights, platform)
    return question_numbers

def data_frame(data: list, file_path: str, selected_options: list):
    output_directory = "Output"
    # Create the output directory if it doesn't exist
    os.makedirs(output_directory, exist_ok=True)
    
    # Get the file name without extension
    file_name = os.path.splitext(os.path.basename(file_path))[0] + ".xlsx"
    output_path = os.path.join(output_directory, file_name)
    
    df = pd.DataFrame(data)
    
    if "Xáo trộn câu hỏi" in selected_options:
        df = df.sample(frac=1)  # frac=1 shuffles all rows randomly
    
    df.to_excel(output_path, index=False)
    # Open the output directory
    os.startfile(output_path)