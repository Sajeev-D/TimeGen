from docx import Document

def create_document(content):
    # Create a new document
    doc = Document()

    # Add content to the document
    doc.add_paragraph(content)

    # Save the document
    doc.save('output.docx')

# Example usage
string_content = "This is the content of the document."
create_document(string_content)

# from docx import Document
# from docx.shared import Cm

# # style.font.size = Pt(12)
# # style.font.name = 'Arial'
# # style.font.bold = True
# # style.font.italic = True
# # style.font.underline = True

# print("Hello World\n")

# text = "DisputeLens. The start of the future.\n"

# print(text)

# doc = Document()

