import os
import base64
import filter
import xmlxexp
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def authenticate():
    creds = None
    if os.path.exists('./.creds/token.json'):
        creds = Credentials.from_authorized_user_file('./.creds/token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('./.creds/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('./.creds/token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def get_messages(service):
    messages = []
    page_token = None

    while True:
        response = service.users().messages().list(userId='me', pageToken=page_token, maxResults=100).execute()
        if 'messages' in response:
            messages.extend(response['messages'])
        page_token = response.get('nextPageToken')
        
        # Если нет следующей страницы, выходим из цикла
        if not page_token:
            break

    return messages

def main():
    creds = authenticate()
    service = build('gmail', 'v1', credentials=creds)
    messages = get_messages(service)

    result = filter.process_messages(messages, service)
    xmlxexp.save_to_excel(result)

if __name__ == "__main__":
    main()
