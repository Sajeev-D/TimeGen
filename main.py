import openai
from docx import Document
from docx.shared import Pt

# Set your OpenAI API key
openai.api_key = 'sk-proj-UGMA2Nv0TPrVSCpl0dngT3BlbkFJYgd9SSZ6dV2avNJLQt0C'


string Prompt = """Look through this email chain, make me a timeline of what was agreed to and when. make it concise. for each email, include the date/time, and who sent it. For the content, only include what was either PROMISED or ACCEPTED."""
string Emails = 

# Make a GPT-3 API call
response = openai.ChatCompletion.create(
  model="gpt-3.5-turbo",
  messages=[
        {"role": "user", "content": Prompt + Emails},
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
    doc.save('thread_summarized.docx')

# Create the document with the GPT-3 response
create_document()
