import os
import docx
import shutil

def doc_to_docx(file_path):
    try: 
        # Split the file path into name and extension
        name, ext = os.path.splitext(os.path.abspath(file_path))
        
        if ext == ".doc":
            # Create a new file path for the copied file with the .docx extension
            new_file_path = name + ".docx"

            # Copy the original file to the new file path
            shutil.copyfile((os.path.abspath(file_path)), new_file_path)
            print("File copied and renamed to", new_file_path)

            # Now, open the copied .docx file
            doc = docx.Document(new_file_path)
        elif ext == ".docx": doc = docx.Document(file_path)
        else: return False
        return doc
    except Exception as e:
        return False

# Example usage:
doc= doc_to_docx(r"Docx\Chương III-k12.doc")
if doc is False: 
    print("CANCLE")
    exit()

for paragraph in doc.paragraphs:
    print(paragraph.text)
    exit()
print("PASS")