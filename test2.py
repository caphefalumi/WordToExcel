from docx import Document
from lxml import etree

def extract_format_text(paragraph):
    # Extracts formatted text (highlighted or bold) from a paragraph.
    format_text = ""
    for run in paragraph.runs:
        print(run.text)
    return format_text
doc = Document(r"Docx\ĐC _ SINH 12.docx")
for paragraph in doc.paragraphs:
    extract_format_text(paragraph)