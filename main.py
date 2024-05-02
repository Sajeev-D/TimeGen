import openai
from docx import Document
from docx.shared import Pt

def read_document():
    # Open the existing document
    doc = Document('output.docx')

    # Read the content of the document
    content = []
    for para in doc.paragraphs:
        content.append(para.text)

    # Join all the paragraphs together
    string_content = ' '.join(content)

    return string_content

# Set your OpenAI API key
openai.api_key = 'sk-proj-UGMA2Nv0TPrVSCpl0dngT3BlbkFJYgd9SSZ6dV2avNJLQt0C'

# Make a GPT-3 API call
response = openai.ChatCompletion.create(
  model="gpt-3.5-turbo",
  messages=[
        {"role": "user", "content": "Write a song about my Malaysian friend Sajeev who likes Nasi Goreng"},
    ]
)

# Get the content from the GPT-3 response
string_content = response['choices'][0]['message']['content']

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

# Create the document with the GPT-3 response
create_document()

# To fix the exceptions issue: do pip install python-docx