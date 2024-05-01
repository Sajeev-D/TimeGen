from docx import Document
from docx.shared import Pt  # Import Pt

# Example usage
string_content = "Dispute Lens is a Billion Dollar Company."

def create_document():
    # Create a new document
    doc = Document()

    # Add content to the document
    doc.add_paragraph(string_content)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(18)  # Use Pt for font size

    # Save the document
    doc.save('output.docx')


