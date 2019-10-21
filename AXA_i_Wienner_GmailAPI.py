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
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

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
    # print()
    # print(user_email)
    # print()





                ###############   AXA   #####################

    try:
        n = 0
        while n < 10:
            today = date.today()
            yesterday = today - timedelta(1)
            query = "newer_than:42d".format(today.strftime('%d/%m/%Y'))
            query01 = "from:faktury_prowizje@axaubezpieczenia.pl"

            results = service.users().messages().list(userId='me', labelIds=['Label_6603011562280603842'],
                                                      maxResults=n + 1, q=query).execute()
            message_id = results['messages'][n]['id']
            msg01 = service.users().messages().get(userId='me', id=message_id).execute()
            msg02 = str(msg01)
            # print(message_id)
            # print(msg01['snippet'])

            if msg02.find('rozliczenie prowizji') > -1:
                print()
                print('AXA ok')

                att_id = ''
                for part in msg01['payload']['parts']:
                    if part['filename']:
                        if 'data' in part['body']:
                            att_id = part['body']['data']
                            # print(att_id)

                        else:
                            att_id = part['body']['attachmentId']
                            # print(att_id)

                # .xls + dopasuj nr wiadomości i id

                get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                       id=att_id).execute()
                get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
                path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2019/AXA prowizja' + '.xls'])
                f = open(path, 'wb')
                f.write(get_att_de)
                f.close()

                # ten fragment zdejmuje hasło z rozliczenia prowizyjnego AXA

                xlApp = Dispatch("Excel.Application")
                xlwb = xlApp.Workbooks.Open('C:\\Users\ROBERT\Desktop\Księgowość\\2019\AXA prowizja.xls', False, False, None,
                                            'PVxCC32%pLkO')
                path = ''.join(['C:\\Users\ROBERT\Desktop\Księgowość\\2019'])
                xlApp.DisplayAlerts = False
                xlwb.SaveAs(path + '\AXA prowizja.xls', FileFormat=-4143, Password='')
                xlApp.DisplayAlerts = True
                xlwb.Close()
                break


            else:
                # print('Brak AXA')
                n += 1


    except:
        print('Brak AXA')






                #############   WIENER   #################

    try:
        n = 0
        while n < 10:
            today = date.today()
            yesterday = today - timedelta(1)
            query = "newer_than:42d".format(today.strftime('%d/%m/%Y'))
            query01 = "from: adres email"

            results = service.users().messages().list(userId='me', labelIds=['Label_7350084330973658333'],
                                                      maxResults=n + 1, q=query).execute()
            message_id = results['messages'][n]['id']
            msg01 = service.users().messages().get(userId='me', id=message_id).execute()
            msg02 = str(msg01)
            # print(message_id)
            # print(msg01['snippet'])

            if msg02.find('prowizji za miesiąc') > -1:

                print('Wiener ok')

                att_id = ''
                for part in msg01['payload']['parts']:
                    if part['filename']:
                        if 'data' in part['body']:
                            att_id = part['body']['data']
                            # print(att_id)

                        else:
                            att_id = part['body']['attachmentId']
                            # print(att_id)



                # .xls + dopasuj nr wiadomości i id

                get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                       id=att_id).execute()
                get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
                path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2019/Wiener prowizja' + '.pdf'])
                f = open(path, 'wb')
                f.write(get_att_de)
                f.close()
                break


            else:
                # print('Brak Wiener')
                n += 1

    except:
        print('Brak Wiener')




        #############   INSLY   #################


    try:
        n = 0
        while n < 10:
            today = date.today()
            yesterday = today - timedelta(1)
            query = "newer_than:42d".format(today.strftime('%d/%m/%Y'))
            query01 = "from: adres email"

            results = service.users().messages().list(userId='me', labelIds=['Label_2969710781820475073'],
                                                      maxResults=n + 1, q=query).execute()
            message_id = results['messages'][n]['id']
            msg01 = service.users().messages().get(userId='me', id=message_id).execute()
            msg02 = str(msg01)
            # print(message_id)
            # print(msg01['snippet'])

            if msg02.find('Faktura') > -1:

                att_id = ''
                for part in msg01['payload']['parts']:
                    # print('part')
                    if 'faktura' in part['filename']:
                        if 'data' in part['body']:
                            att_id = part['body']['data']

                        else:
                            att_id = part['body']['attachmentId']


                get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                       id=att_id).execute()
                get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
                path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2019/Insly faktura' + '.pdf'])
                f = open(path, 'wb')
                f.write(get_att_de)
                f.close()
                print('Insly ok')
                break


        else:
            # print('Brak Insly')
            n += 1

    except:
        print('Brak Insly')






            #######################  ORANGE MOBILNE ########################

    try:
        n = 0
        while n < 10:
            today = date.today()
            yesterday = today - timedelta(1)
            query = "newer_than:42d".format(today.strftime('%d/%m/%Y'))
            query01 = "from: adres email"

            results = service.users().messages().list(userId='me', labelIds=['Label_7521852298094424071'],
                                                      maxResults=n + 1, q=query).execute()
            message_id = results['messages'][n]['id']
            msg01 = service.users().messages().get(userId='me', id=message_id).execute()
            msg02 = str(msg01)
            # print(message_id)
            # print(msg01['snippet'])

            if msg02.find('e-fakturę za usługi mobilne') > -1:

                att_id = ''
                for part in msg01['payload']['parts']:
                    # print('part')
                    if re.search('(?=.*FAKTURA)(?!.*xml).*', part['filename']):
                        if 'data' in part['body']:
                            att_id = part['body']['data']

                        else:
                            att_id = part['body']['attachmentId']

                get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                       id=att_id).execute()
                get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
                path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2019/Orange faktura usł mobilne' + '.pdf'])
                f = open(path, 'wb')
                f.write(get_att_de)
                f.close()
                print('Orange mobilne ok')
                break

            else:
                n += 1

    except:
        print('Brak Orange usł mobilne')






            #######################  ORANGE STACJONARNE ########################

    try:
        n = 0
        while n < 10:
            today = date.today()
            yesterday = today - timedelta(1)
            query = "newer_than:42d".format(today.strftime('%d/%m/%Y'))
            query01 = "from: adres email"

            results1 = service.users().messages().list(userId='me', labelIds=['Label_7521852298094424071'],
                                                      maxResults=n + 1, q=query).execute()
            message_id = results1['messages'][n]['id']
            msg03 = service.users().messages().get(userId='me', id=message_id).execute()
            msg04 = str(msg03)

            if msg04.find('e-faktura Orange') > -1:

                att_id = ''
                for part in msg03['payload']['parts']:
                    # print('part')
                    if re.search('[0-9](?!.*xml).*', part['filename']):
                        if 'data' in part['body']:
                            att_id = part['body']['data']

                        else:
                            att_id = part['body']['attachmentId']

                get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                       id=att_id).execute()
                get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
                path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2019/Orange faktura usł stacjonarne' + '.pdf'])
                f = open(path, 'wb')
                f.write(get_att_de)
                f.close()
                print('Orange stacjonarne ok')
                break

            else:
                n += 1


    except:
        print('Brak Orange usł stacjonarne')






        ####################### AWS ########################


    try:
        n = 0
        while n < 10:
            today = date.today()
            yesterday = today - timedelta(1)
            query = "newer_than:39d".format(today.strftime('%d/%m/%Y'))
            query01 = "from: adres email"

            results = service.users().messages().list(userId='me', labelIds=['Label_3955391925081514655'],
                                                      maxResults=n + 1, q=query).execute()
            message_id = results['messages'][n]['id']
            msg01 = service.users().messages().get(userId='me', id=message_id).execute()
            msg02 = str(msg01)
            # print(message_id)
            # print(msg01['snippet'])

            if msg02.find('VAT Invoice(s) available') > -1:

                print('AWS ok')

                att_id = ''
                for part in msg01['payload']['parts']:
                    if part['filename']:
                        if 'data' in part['body']:
                            att_id = part['body']['data']
                            # print(att_id)

                        else:
                            att_id = part['body']['attachmentId']
                            # print(att_id)


                get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                       id=att_id).execute()
                get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
                path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2019/AWS faktura' + '.pdf'])
                f = open(path, 'wb')
                f.write(get_att_de)
                f.close()
                break


            else:
                # print('Brak Wiener')
                n += 1

    except:
        print('Brak AWS')



if __name__ == '__main__':
    main()
