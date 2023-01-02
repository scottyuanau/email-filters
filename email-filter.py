from __future__ import print_function

import os.path
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.settings.basic']


def scottpersonal():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    path = '/Users/scott/Learning/Coding/Python3/email-filter/credentials/personal-account'
    if os.path.exists(f'{path}/token.json'):
        creds = Credentials.from_authorized_user_file(f'{path}/token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                f'{path}/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(f'{path}/token.json', 'w') as token:
            token.write(creds.to_json())
    df = pd.read_excel('email-blacklist.xlsx')

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        existingfilters = service.users().settings().filters().list(userId='me').execute()
        # list filters
        emails = df['Email Black List'].tolist()
        for email in emails:
            filter_content = {
                'criteria': {
                    'from': email
                },
                'action': {
                    'addLabelIds': ['TRASH'],
                    'removeLabelIds': ['INBOX']
                }
            }
            addFilter = service.users().settings().filters().create(userId='me', body=filter_content).execute()
            print(f'Created filter with id:{addFilter.get("id")}')


    except HttpError as error:
        print(f'An error occurred: {error}')


if __name__ == '__main__':
    scottpersonal()