import docx
import pandas as pd
import os
import win32com.client
from docx.enum.text import WD_COLOR_INDEX as color
import docx.enum.text

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
