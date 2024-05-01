from docx import Document

def create_document(content):
    # Create a new document
    doc = Document()

    # Add content to the document
    doc.add_paragraph(content)

    # Save the document
    doc.save('output.docx')

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = 18

# Example usage
string_content = "Dispute Lens is a Billion Dollar Company."
create_document(string_content)