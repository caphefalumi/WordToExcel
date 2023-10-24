import os, shutil, pypandoc, docx
import win32com.client as win32
from win32com.client import constants
import re
file_path = r"C:\Users\Toan\WordToExcel\Docx\Chương III-k12.doc"
def convert(file_path):
    # Load the DOCX document
    pypandoc.convert_file(file_path, 'plain', outputfile='temp.txt')
    document = docx.Document()
    myfile = open(f'wordToExcelConvertTemp.txt', 'r', encoding='utf-8').read()
    document.add_paragraph(myfile)
    document.save('wordToExcelConvertTemp.docx')
    os.remove('wordToExcelConvertTemp.txt')
    return os.path.abspath('wordToExcelConvertTemp.docx') 
convert(file_path)