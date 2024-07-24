from docx import Document

from transformers import pipeline

def read_docx(file_path):
    """
    Read the content of a DOCX file and return it as a plain text string.
    """
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def summarize_text(text):
    """
    Summarize the provided text using a pre-trained model.
    """
    # Load a pre-trained summarization model
    summarizer = pipeline("summarization", model="facebook/bart-large-cnn")
    summary = summarizer(text, max_length=130, min_length=30, do_sample=False)
    return summary[0]['summary_text']

def main():
    # Path to your .docx file
    docx_file = 'timeline.docx'
    # Path to save the summary text
    summary_file = 'summary.txt'
    
    # Read the text from the document
    document_text = read_docx(docx_file)
    # Generate a summary of the document
    summary = summarize_text(document_text)
    
    # Print the summary to the console
    print("Document Summary:")
    print(summary)

    # Save the summary to a text file
    with open(summary_file, 'w') as file:
        file.write(summary)

if __name__ == "__main__":
    main()

# This section summarises the content of a DOCX file using a pre-trained model from the Hugging Face Transformers library. The script reads the content of a DOCX file, generates a summary of the text, and saves the summary to a new text file. The `read_docx` function reads the content of a DOCX file and returns it as a plain text string. The `summarize_text` function uses a pre-trained summarization model to generate a summary of the provided text. The `main` function reads the content of a DOCX file, generates a summary of the text, prints the summary to the console, and saves the summary to a text file. To use this script, you need to provide the path to the input DOCX file and specify the path to save the summary text. You can adjust the summarization parameters such as `max_length` and `min_length` to control the length of the generated summary. The script uses the `facebook/bart-large-cnn` model for summarization, but you can replace it with other pre-trained models available in the Transformers library.
# Was contributed by Iftikher Zaman
