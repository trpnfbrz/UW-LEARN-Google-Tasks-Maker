import imaplib
import email
import yaml

import html2text
h = html2text.HTML2Text() #create instance of HTML2Text
h.ignore_links = True

import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

CLIENT_FILE = 'client_secret.json'
TASKS_SCOPES = 'https://www.googleapis.com/auth/tasks'
GMAIL_SCOPES = 'https://mail.google.com/'

MY_TASKS = 'MDA3Njg0NTA5OTg0MjgzODMzMjc6MDow'
DAILY_TO_DO = 'VFBmTWROd2lsa1dKSXl5Rw'
LEARN_NOTIFS = 'alF6R3c4eHdhdEdOSHc2OQ'

creds = None

if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', [TASKS_SCOPES, GMAIL_SCOPES])

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else: 
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_FILE, [TASKS_SCOPES, GMAIL_SCOPES])
        creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())
    
# Create service object instances
tasks_service = build('tasks', 'v1', credentials=creds)
gmail_service = build('gmail', 'v1', credentials=creds)

# Call the Tasks API to retrieve all task lists
results = tasks_service.tasklists().list().execute()
items = results.get('items', [])

if not items:
    print('No task lists found.')
else:
    print('Your task lists:')
    for item in items:
        print(f"Title: {item['title']}, ID: {item['id']}")

def create_google_task(service, tasklist_id, title, notes):
    """Creates a new task in a specified Google Tasks list."""
    task = {
        'title': title,
        'notes': notes
    }
    try:
        results = service.tasks().insert(tasklist=tasklist_id, body=task).execute()
        print(f"Task created: {results['title']}")
    except HttpError as error:
        print(f"An error occurred: {error}")

def search_messages(service, query):
    result = service.users().messages().list(userId='me',q=query).execute()
    messages = [ ]
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me',q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages

def archive_mail(service, query): # Function to archive read emails
    messages_to_mark = search_messages(service, query)
    print(f"Archived emails: {len(messages_to_mark)}")
    return service.users().messages().batchModify(
      userId='me',
      body={
          'ids': [ msg['id'] for msg in messages_to_mark ],
          'removeLabelIds': ['INBOX']
      }
    ).execute()

def inbox_scraper(): # Function to scrape emails
    with open("credentials.yml") as f:
        content = f.read()
        
    my_credentials = yaml.load(content, Loader=yaml.FullLoader)

    user, password = my_credentials["user"], my_credentials["password"]

    imap_url = 'imap.gmail.com'

    my_mail = imaplib.IMAP4_SSL(imap_url)

    my_mail.login(user, password)

    my_mail.select('Inbox')

    _, data = my_mail.search(None, 'FROM', 'learnhlp@uwaterloo.ca')

    mail_id_list = data[0].split()

    msgs = []
    for num in mail_id_list:
        typ, data = my_mail.fetch(num, '(RFC822)') #RFC822 returns whole message
        msgs.append(data)

    for msg in msgs[::-1]: 
        for response_part in msg:
            if type(response_part) is tuple:
                my_msg=email.message_from_bytes((response_part[1]))
                # Print Mail
                # print("_________________________________________")
                # print ("subj:", my_msg['subject'])
                # print ("from:", my_msg['from'])
                # print ("body:")
                for part in my_msg.walk():
                    # print(part.get_content_type())
                    # print(h.handle(part.get_payload()))
                    print('')

        try: # Pass the instance to the function
            create_google_task(tasks_service, LEARN_NOTIFS, my_msg['subject'], h.handle(part.get_payload()))
        except Exception as e:
            print(f"An error occurred: {e}")

inbox_scraper()

try: # Archive read notification
    archive_mail(gmail_service, 'learnhlp@uwaterloo.ca')
except Exception as e:
    print(f"An error occurred: {e}")


