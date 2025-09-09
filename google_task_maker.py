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

CLIENT_FILE = 'client_secret_577299332042-jc7q7nibi10kcbr2o4rhllhio6uc273m.apps.googleusercontent.com'
API_NAME = 'tasks'
API_VERSION = 'v1'
SCOPES = ['https://www.googleapis.com/auth/tasks']

creds = None

if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else: 
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

def inbox_scraper(): #function to scrape emails
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
                print("_________________________________________")
                print ("subj:", my_msg['subject'])
                print ("from:", my_msg['from'])
                print ("body:")
                for part in my_msg.walk():
                    #print(part.get_content_type())
                    print(h.handle(part.get_payload()))

inbox_scraper()
