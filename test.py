import docx
doc = docx.Document(r"C:\Users\Toan\WordToExcel\Docx\Toàn Đặng Chuột - ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA KÌ I MÔN GDQP-AN 12 (1) (1).docx")
for paragraph in doc.paragraphs:
    text = paragraph.text.strip()
    if text.startswith("Câu"):
        print(text)