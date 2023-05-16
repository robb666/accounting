
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import date, timedelta
import base64
from win32com.client import Dispatch
import re


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://mail.google.com/']


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    # project_path = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno\dist\\'
    project_path = r'C:\Users\PipBoy3000\Desktop\IT\projekty\accounting\\'
    if os.path.exists(project_path + 'token.pickle'):
        with open(project_path + 'token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(project_path + 'credentials.json', SCOPES)
            creds = flow.run_local_server()

        # Save the credentials for the next run
        with open(project_path + 'token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds)

    # # """CHECK"""
    #
    # service = build('gmail', 'v1', credentials=creds)
    #
    # # Call the Gmail API
    # results = service.users().labels().list(userId='me').execute()
    # labels = results.get('labels', [])
    #
    # if not labels:
    #     print('No labels found.')
    # else:
    #     print('Labels:')
    #     for label in labels:
    #         print(label['name'] + ' ' + label['id'])
    #
    # user_profile = service.users().getProfile(userId='me').execute()
    # user_email = user_profile['emailAddress']
    # print(user_email)


def zallianz():
    label = {'zallianz': 'Label_3251381808219322746'}
    today = date.today()
    query = "newer_than:1d".format(today.strftime('%d/%m/%Y'))
    results = service.users().messages().list(userId='me', labelIds=[label['zallianz']], maxResults=1, q=query).execute()
    message_id = results['messages'][0]['id']
    msg = service.users().messages().get(userId='me', id=message_id).execute()
    tiktok = re.search('jednorazowy (\d+)', msg['snippet'])
    if tiktok:
        return tiktok.group(1)


def zsanpl():
    label = {'zallianz/zsanpl': 'Label_7938073158094859915'}
    today = date.today()
    query = "newer_than:1d".format(today.strftime('%d/%m/%Y'))
    results = service.users().messages().list(userId='me',
                                              labelIds=[label['zallianz/zsanpl']],
                                              maxResults=1,
                                              q=query).execute()
    message_id = results['messages'][0]['id']
    msg = service.users().messages().get(userId='me', id=message_id).execute()
    zsan = re.search('od: (\d{3}-\d{3})', msg['snippet'])
    if zsan:
        return zsan.group(1)


service = main()

# next_month_path = 'C:\\Users\\PipBoy3000\\Desktop\\Księgowość\\10.2028\\'
# email(next_month_path)
