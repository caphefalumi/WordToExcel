import re
def is_option(paragraph):
    # Checks if a paragraph starts with an option (A., B., C., D.).
    
    return paragraph.startswith(("A.", "B.", "C.", "D.","a.", "b.", "c.", "d."))

# Helper function to split options that are on the same line
def split_options(text):
    # Splits options that are on the same line into a list.
    return re.split(r'\s+(?=[a-dA-D]\.)', text)
data = []
current_options = ["Anh, Pháp, Mỹ.", "B.Đức, Italia, Nhật","C.Anh, Pháp, Liên Xô"]
data.append({
    'Question Type': "Multiple Choice",
    'Option 1': current_options[0] if len(current_options) > 0 else "",
    'Option 2': current_options[1] if len(current_options) > 1 else "",
    'Option 3': current_options[2] if len(current_options) > 2 else "",
    'Option 4': current_options[3] if len(current_options) > 3 else "",
    'Time in seconds': 30,
})

for i in data:
    print(i)