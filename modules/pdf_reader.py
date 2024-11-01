from pypdf import PdfReader

def read_pdf(file_path):
    lines = []
    reader = PdfReader(file_path)
    
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text = page.extract_text()
        if text:
           lines.extend(text.split('\n'))
    
    lines = [item for item in lines if item not in [" ", "  "]]
    formatted_lines = []
    
    for line in lines:
        formatted_lines.append({'text': line, 'style': 'Normal'})

    return formatted_lines
