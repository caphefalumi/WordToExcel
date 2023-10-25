import docx
import re

doc = docx.Document(r"C:\Users\Toan\WordToExcel\Docx\Dap an.docx")
count = 0
for paragraph in doc.paragraphs:
    text = paragraph.text.strip()
    text = re.sub("\x0B", "\x0A", text)
    count +=1
    print(text, count)