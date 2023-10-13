import os, subprocess, re
import time
import win32com.client as win32
from win32com.client import constants

def close_word(file_name):
    if os.path.exists(file_name):
        # Closes an Excel application if it is open.
        try:
            subprocess.call("TASKKILL /F /IM WINWORD.EXE", shell=True)
        except subprocess.CalledProcessError: 
            pass
abs_path = os.path.abspath(r"Docx\Dap an.docx")
def save_as_docx(path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate ()

    # Rename path with .docx
    abs_path = os.path.abspath(path)
    ext = os.path.splitext(abs_path)[1]
    if ext == ".doc":
        new_file_abs = re.sub(r'\.\w+$', '.docx', abs_path)
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        return new_file_abs
    if ext == ".docx":
        new_file_abs_1 = re.sub(r'\.\w+$', '.doc', abs_path)
        
        # Save as .doc
        word.ActiveDocument.SaveAs(new_file_abs_1, FileFormat=constants.wdFormatDocument)
        
        # Close the document
        doc.Close(False)
        
        # Reopen the saved .doc file
        doc = word.Documents.Open(new_file_abs_1)
        doc.Activate()
        
        # Rename it back to .docx
        new_file_abs_2 = re.sub(r'\.\w+$', '1.docx', new_file_abs_1)
        # Save as .docx
        word.ActiveDocument.SaveAs(new_file_abs_2, FileFormat=constants.wdFormatXMLDocument)
        os.remove(new_file_abs_1)
        doc.Close(False)
        return new_file_abs_2
path = save_as_docx(abs_path)



if os.path.exists(path):
    print("SUCCESS")
subprocess.call("TASKKILL /F /IM WINWORD.EXE", shell=True)