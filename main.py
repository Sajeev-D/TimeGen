########################### Start of imports ###########################

import requests
import msal
import openai
import asyncio
import win32com.client # pip install pywin32
import datetime
import os
import glob

from docx import Document
from docx.shared import Pt
from oauthlib.oauth2 import WebApplicationClient
from msgraph import GraphServiceClient
#from azure.identity import ClientSecret
from azure.identity import ClientSecretCredential
from pathlib import Path


########################### End of imports ###########################
#
#
#
#
########################### Start of helper functions ###########################

# read the email thread from emails_raw.docx and return the content as a string
def read_document():
    # Open the existing document
    doc = Document('emails_raw.docx')

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

def getEmails():
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
        if message.ReceivedTime.date() < datetime.date(2024, 5, 3):
            print("Emails after the specified date have been saved to the output folder.")
            break

        #Get the email subject
        subject = message.Subject
        #Get the email body
        body = message.Body
        #Get the received date and time
        received_time = message.ReceivedTime
        #Save the email body and received time to a text file
        with open(output_dir / f'{num}.txt', 'w', encoding='utf-8') as f:
            f.write(f"Received Time: {received_time}\n\n")
            f.write(body)   
        num += 1

    return

        
########################### End of helper functions ###########################
#
#
#
#
########################### Start of main function ###########################

def main():    
    ## Creates a timeline of the email thread
    # firstPrompt = "Look through this email chain, make me a timeline of what was agreed to and when. make it concise. for each email, include the date/time, and who sent it. For the content, only include what was either PROMISED or ACCEPTED."
    # Emails = read_document()
    # fullPrompt = firstPrompt + "\n" + Emails
    # # string_content = getGPT3Response(fullPrompt)
    # string_content = "This is a test string."
    # create_document(string_content)

    print('Getting emails...\n')
    getEmails()
    print('\nEmails have been saved to the output folder.')
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
#
#
#
#
########################### Code Dump ###########################

# #initialize the variables
# client_id = '8a23d535-8a92-469b-a11b-d6fa336eea9a'
# client_secret = '8e366d84-4076-460e-8f92-a87b680f93da'
# tenant_id = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a'
# authority = f'https://login.microsoftonline.com/f8cdef31-a31e-4b4a-93e4-5f571e91255a'
# scope = ['https://graph.microsoft.com/User.Read','https://graph.microsoft.com/Mail.Read' ] #done

# # Initialize the OAuth client
# client = WebApplicationClient(client_id)
# redirect_uri = 'http://localhost:8000/callback'

# # Step 1: Redirect the user to the provider's authorization page
# authorization_url = client.prepare_request_uri(
#     "https://provider.com/oauth/authorize",
#     redirect_uri=redirect_uri,
#     scope=["mail.read"],
# )
# # Redirect the user to authorization_url

# # Step 2: User is redirected back to your application, capture the code
# code = request.args.get("code")

# # Step 3: Exchange the code for a token
# token_url, headers, body = client.prepare_token_request(
#     "https://provider.com/oauth/token",
#     authorization_response=request.url,
#     redirect_url=request.base_url,
#     code=code,
# )
# token_response = requests.post(
#     token_url,
#     headers=headers,
#     data=body,
#     auth=(client_id, client_secret),
# )

# # Parse the tokens!
# client.parse_request_body_response(json.dumps(token_response.json()))

# app = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)

# result = app.acquire_token_for_client(scope)

# access_token = None

# response = None

# if 'access_token' in result:
#     access_token = result['access_token']
#     graph_endpoint = 'https://graph.microsoft.com/v1.0/me/messages'
#     headers = {'Authorization': 'Bearer ' + access_token}
#     response = requests.get(graph_endpoint, headers=headers)
#     print(response.json())
# else:
#     print(f'Could not acquire token: {result}')

# async def graph_API_call():
#     credential = ClientSecretCredential(
#         tenant_id="f8cdef31-a31e-4b4a-93e4-5f571e91255a",
#         client_id="8a23d535-8a92-469b-a11b-d6fa336eea9a",
#         client_secret="8e366d84-4076-460e-8f92-a87b680f93da",
#     )

#     client = GraphClient(credential=credential, scopes=['https://graph.microsoft.com/.default'])

#     message_id = 'message-id'  # replace with your message id
#     endpoint = f'https://graph.microsoft.com/v1.0/me/messages/{message_id}'

#     headers = {
#         "Prefer": "outlook.body-content-type=\"text\""
#     }

#     response = await client.get(endpoint, headers=headers)
#     email_content = response.json()

#     return email_content

# async def list_messages():
#     credential = ClientSecretCredential(
#         tenant_id="f8cdef31-a31e-4b4a-93e4-5f571e91255a",
#         client_id="8a23d535-8a92-469b-a11b-d6fa336eea9a",
#         client_secret="8e366d84-4076-460e-8f92-a87b680f93da",
#     )

#     loop = asyncio.get_event_loop()
#     token = await loop.run_in_executor(None, lambda: credential.get_token('https://graph.microsoft.com/.default'))

#     url = 'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages'

#     headers = {
#         'Authorization': 'Bearer ' + token.token
#     }

#     response = requests.get(url, headers=headers)
#     messages = response.json()

#     for message in messages['value']:
#         print(message['id'])

########################### End of Code Dump ###########################