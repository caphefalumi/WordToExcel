import win32com.client as win32
import shutil

# Create a Word Application object
word = win32.gencache.EnsureDispatch("Word.Application")

# Define the paths for the original and copy documents
original_file_path = r'C:\Users\Toan\WordToExcel\Docx\Toàn Đặng Chuột - ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA KÌ I MÔN GDQP-AN 12 (1) (1).docx'
copy_file_path = r'C:\Users\Toan\WordToExcel\Docx\TEST.docx'

# Open the original document
doc = word.Documents.Open(original_file_path)

# Create a copy of the original document
shutil.copy(original_file_path, copy_file_path)

# Open the copied document
copy_doc = word.Documents.Open(copy_file_path)

# Define your VBA macro code
vba_code = """
Public Sub Test()

  Dim oRng As Range
  Dim CC   As ContentControl
  Dim LC   As Integer
  Dim LRCC As Integer
  Dim LTCC As Integer
  Dim LE   As Boolean

Set oRng = ActiveDocument.Content
LTCC = LTCC + oRng.ContentControls.Count
For LC = oRng.ContentControls.Count To 1 Step -1
 
Set CC = oRng.ContentControls(LC)
If CC.LockContentControl = True Then
    CC.LockContentControl = False
End If
CC.Delete
If Not LE Then
    LRCC = LRCC + 1
    End If
    LE = False
Next
End Sub
"""

# Insert the macro code into the copied Word document
copy_doc.VBProject.VBComponents.Add(1).CodeModule.AddFromString(vba_code)

# Save and close the copied document
copy_doc.Save()
copy_doc.Close()

# Close the original document
doc.Close()

# Release the Word Application object
word.Quit()
