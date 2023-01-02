from __future__ import print_function

import os.path
import time

import pandas as pd
import schedule

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.settings.basic']
blacklist = '/Users/scott/Library/CloudStorage/OneDrive-Personal/email-blacklist.xlsx'

# def blockemailmaster():
#     """Shows basic usage of the Gmail API.
#     Lists the user's Gmail labels.
#     """
#     creds = None
#     accountname = 'Block Email Account'
#     # The file token.json stores the user's access and refresh tokens, and is
#     # created automatically when the authorization flow completes for the first
#     # time.
#     path = '/Users/scott/Learning/Coding/Python3/email-filter/credentials/block-email-account'
#     #modify the path to the correct credential folder
#     MAILSCOPES = 'https://www.googleapis.com/auth/gmail.modify'
#     if os.path.exists(f'{path}/token.json'):
#         creds = Credentials.from_authorized_user_file(f'{path}/token.json', MAILSCOPES)
#     # If there are no (valid) credentials available, let the user log in.
#     if not creds or not creds.valid:
#         if creds and creds.expired and creds.refresh_token:
#             creds.refresh(Request())
#         else:
#             flow = InstalledAppFlow.from_client_secrets_file(
#                 f'{path}/credentials.json', MAILSCOPES)
#             creds = flow.run_local_server(port=0)
#         # Save the credentials for the next run
#         with open(f'{path}/token.json', 'w') as token:
#             token.write(creds.to_json())
#     df = pd.read_excel('email-blacklist.xlsx')
#
#     try:
#         # Call the Gmail API
#         service = build('gmail', 'v1', credentials=creds)
#         reademails = service.users().messages().list(userId='me').execute()
#         print(reademails)
#
#
#
#     except HttpError as error:
#         print(f'An error occurred: {error}')
def scottpersonal():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    accountname = 'Scott Personal Account'
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
    df = pd.read_excel(blacklist)

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        existingfilters = service.users().settings().filters().list(userId='me').execute()
        # list filters
        existingfilter = []
        for item in existingfilters['filter']:
            if 'from' in item['criteria'].keys():
                existingfilter.append(item['criteria']['from'])
        #check existing filters

        emails = df['Email Black List'].tolist()
        emailadded = False
        for email in emails:
            if email not in existingfilter:
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
                emailadded = True
                print(f'{email} added to {accountname} blacklist.')

        if not emailadded:
            print(f'No new email added to {accountname}.')

    except HttpError as error:
        print(f'An error occurred: {error}')
def roaster():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    accountname = 'Roaster Account'
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    path = '/Users/scott/Learning/Coding/Python3/email-filter/credentials/roaster-account'
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
    df = pd.read_excel(blacklist)

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        existingfilters = service.users().settings().filters().list(userId='me').execute()
        # list filters
        existingfilter = []
        for item in existingfilters['filter']:
            if 'from' in item['criteria'].keys():
                existingfilter.append(item['criteria']['from'])
        #check existing filters

        emails = df['Email Black List'].tolist()
        emailadded = False
        for email in emails:
            if email not in existingfilter:
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
                emailadded = True
                print(f'{email} added to {accountname} blacklist.')

        if not emailadded:
            print(f'No new email added to {accountname}.')

    except HttpError as error:
        print(f'An error occurred: {error}')
def scottwork():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    accountname = 'Scott Work Account'
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    path = '/Users/scott/Learning/Coding/Python3/email-filter/credentials/scott-work-account'
    #modify the path to the correct credential folder

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
    df = pd.read_excel(blacklist)

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        existingfilters = service.users().settings().filters().list(userId='me').execute()
        # list filters
        existingfilter = []
        for item in existingfilters['filter']:
            if 'from' in item['criteria'].keys():
                existingfilter.append(item['criteria']['from'])
        #check existing filters

        emails = df['Email Black List'].tolist()
        emailadded = False
        for email in emails:
            if email not in existingfilter:
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
                emailadded = True
                print(f'{email} added to {accountname} blacklist.')

        if not emailadded:
            print(f'No new email added to {accountname}.')

    except HttpError as error:
        print(f'An error occurred: {error}')

def cafeaccount():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    accountname = 'Cafe Account'
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    path = '/Users/scott/Learning/Coding/Python3/email-filter/credentials/cafe-account'
    #modify the path to the correct credential folder

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
    df = pd.read_excel(blacklist)

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        existingfilters = service.users().settings().filters().list(userId='me').execute()
        # list filters
        existingfilter = []
        for item in existingfilters['filter']:
            if 'from' in item['criteria'].keys():
                existingfilter.append(item['criteria']['from'])
        #check existing filters

        emails = df['Email Black List'].tolist()
        emailadded = False
        for email in emails:
            if email not in existingfilter:
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
                emailadded = True
                print(f'{email} added to {accountname} blacklist.')

        if not emailadded:
            print(f'No new email added to {accountname}.')

    except HttpError as error:
        print(f'An error occurred: {error}')

def autojob():
    # blockemailmaster()
    scottpersonal()
    roaster()
    scottwork()
    cafeaccount()

#auto run the job
# schedule.every(1).hour.do(autojob)

if __name__ == '__main__':
#auto run the script
    # while True:
    #     schedule.run_pending()
    #     time.sleep(1)
#if want to manually run, use this instead
    autojob()