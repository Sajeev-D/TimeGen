########################### Start of imports ###########################

import requests
import msal
import openai
import asyncio
import win32com.client # pip install pywin32
import datetime
import os
import glob
import re

from docx import Document
from docx.shared import Pt
from oauthlib.oauth2 import WebApplicationClient
from msgraph import GraphServiceClient
#from azure.identity import ClientSecret
from azure.identity import ClientSecretCredential
from pathlib import Path
from datetime import date


########################### End of imports ###########################
#
#
#
#
########################### Start of helper functions ###########################

def txt_to_docx(txt_file, docx_file):
    # Create a new Document
    doc = Document()
    # Open the .txt file
    with open(txt_file, 'r', encoding='utf-8') as file:
        # Read the file
        text = file.read()
        # Add the text to the Document
        doc.add_paragraph(text)
        # Save the Document
        doc.save(docx_file)

    return

# read the email thread from emails_raw.docx and return the content as a string
def read_document(docx_file_name):
    # Open the existing document
    doc = Document(docx_file_name)

    # Read the content of the document
    content = []
    for para in doc.paragraphs:
        content.append(para.text)

    # Join all the paragraphs together
    string_content = ' '.join(content)

    return string_content

# Create a new document with any string argument
def create_document(string_content):
    # Create a new document
    doc = Document()
    # Add content to the document
    doc.add_paragraph(string_content)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)  # Use Pt for font size

    # Save the document
    doc.save('email_thread.docx')

    return

def getGPT3Response(prompt):
    # Set your OpenAI API key
    openai.api_key = 'sk-proj-UGMA2Nv0TPrVSCpl0dngT3BlbkFJYgd9SSZ6dV2avNJLQt0C'
    # Make a GPT-3.5 API call
    response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
            {"role": "user", "content": prompt},
        ]
    )
    # Get the content from the GPT-3 response
    string_content = response['choices'][0]['message']['content']

    return string_content

def getEmails(givenDateFrom, givenSubject, emailWith, yourEmail):
    # Convert the given date to a datetime object
    # if isinstance(givenDateFrom, str):
    #     givenDateFrom = datetime.strptime(givenDateFrom, '%Y-%m-%d').date()

    # Create output folder
    output_dir = Path.cwd() / 'output'
    output_dir.mkdir(parents=True,exist_ok=True)

    # Empty the output folder
    files = glob.glob(str(output_dir / '*'))
    for f in files:
        os.remove(f)

    #Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    #Connect to folder
    inbox = outlook.GetDefaultFolder(6) #6 is the inbox folder ID

    #Get all emails in the inbox
    messages = inbox.Items

    #Sort the emails by ReceivedTime, in descending order
    messages.Sort("[ReceivedTime]", True)

    num = 1

    for message in messages:
         # Check if the item is an email message
        if hasattr(message, "MessageClass") and message.MessageClass.startswith("IPM.Note"):
            if message.ReceivedTime.date() < givenDateFrom:
                print("\nEmails after the specified date have been saved to the output folder.")
                break

            # Get recipients email address
            to_email_address = "default"
            body2 = message.Body

            # Get the line that starts with 'To:'
            to_line = next((line for line in body2.split('\n') if line.startswith('To:')), None)

            if to_line is not None:
                # Use a regular expression to find the email address
                match = re.search(r'<(.+)>', to_line)
                if match:
                    to_email_address = match.group(1)
                    print(f"Receivers Email Address is: {to_email_address}")
            else:
                print("No 'To:' line found in the email body.")
            
            # Select all after a certain date
            if givenSubject == "default" and emailWith == "default" and yourEmail == "default":  
                print('Entered if statement 1\n')              
                #Get the email subject
                subject = message.Subject
                #Get the email body
                body = message.Body
                
                #Get the received date and time
                received_time = message.ReceivedTime
                #Save the email body and received time to a text file
                with open(output_dir / 'all_emails.txt', 'a', encoding='utf-8') as f:
                    f.write(f"Email: {num}\n")
                    f.write(f"Subject: {subject}\n\n")
                    f.write(f"Received Time: {received_time}\n\n")
                    f.write(body)
                    f.write("\n----------------------\n")
                num += 1

            # Select all emails with a specific subject
            if givenSubject != "default" and emailWith == "default":
                if message.Subject == givenSubject or message.Subject == "Re: " + givenSubject:
                    print('Entered if statement 2\n')  
                    #Get the email subject
                    subject = message.Subject
                    #Get the email body
                    body = message.Body

                    print(body)
                    #Get the received date and time
                    received_time = message.ReceivedTime
                    #Save the email body and received time to a text file
                    with open(output_dir / 'all_emails.txt', 'a', encoding='utf-8') as f:
                        f.write(f"Email: {num}\n")
                        f.write(f"Subject: {subject}\n\n")
                        f.write(f"Received Time: {received_time}\n\n")
                        f.write(body)
                        f.write("\n----------------------\n")
                    num += 1
                    if message.Subject == "Re: " + givenSubject:
                        break

            if givenSubject == "default" and emailWith != "default" and yourEmail != "default":
                print('Entered if statement 3\n')  
         
                if emailWith == to_email_address:                    
                    print('Entered the second level function\n\n')
                    subject = message.Subject
                    #Get the email body
                    body = message.Body
                    #Get the received date and time
                    received_time = message.ReceivedTime
                    #Save the email body and received time to a text file
                    with open(output_dir / 'all_emails.txt', 'a', encoding='utf-8') as f:
                        f.write(f"Email: {num}\n")
                        f.write(f"Subject: {subject}\n\n")
                        f.write(f"Received Time: {received_time}\n\n")
                        f.write(body)
                        f.write("\n----------------------\n")
                    num += 1
    return

def is_directory_empty(path):
    return len(os.listdir(path)) == 0
########################### End of helper functions ###########################
#
#
#
#
########################### Start of main function ###########################

def main():    
    if os.path.exists('raw_emails.docx'):
        # Delete the file
        os.remove('raw_emails.docx')

    print('Getting emails...\n')

    # Set the date and subject of the emails to be retrieved
    dateFrom = date(2024, 5, 3)
    subject = 'default'
    emailW = "daniel.parsons@mail.utoronto.ca"
    yourE = "sajeev.debnath@mail.utoronto.ca"

    getEmails(dateFrom, subject, emailW, yourE)
    print('\nEmails have been saved to the output folder.')
    
    firstPrompt = "Look through this email chain, make me a timeline of what was agreed to and when. make it concise. for each email, include the date/time, and who sent it. For the content, only include what was either PROMISED or ACCEPTED."
    
    if is_directory_empty('output'):
        print(f"The directory {'output'} is empty.")

    else:
        print(f"The directory {'output'} is not empty.")
        txt_to_docx('output/all_emails.txt', 'raw_emails.docx')

        Emails = read_document('raw_emails.docx')
        os.remove('raw_emails.docx')

        fullPrompt = firstPrompt + "\n" + Emails
        string_content = getGPT3Response(fullPrompt)

        create_document(string_content)

    return

########################### End of main function ###########################
#
#
#
#
########################### Start of execution ###########################
print('Starting DisputeLens...\n')
main()
print('\nDisputeLens has finished running!')
########################### End of execution ###########################