
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
    project_path = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno\dist\\'
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
    # print()
    # print(user_email)
    # print()


def labels(service):

    labels = {'AXA': 'Label_6603011562280603842',
              'Wiener': 'Label_7350084330973658333',
              'Insly': 'Label_2969710781820475073',
              'Orange stac': 'Label_7521852298094424071',
              'Orange mob': 'Label_7521852298094424071',
              'TUW': 'Label_7255175017814621709',
              'TUZ': 'Label_1453748131451092882',
              'A-Z': 'Label_4747893535910550011',
              'AWS': 'Label_3955391925081514655',
              'Euroins': 'Label_2774382001212357899'}

    today = date.today()
    query = "newer_than:40d".format(today.strftime('%d/%m/%Y'))
    query01 = "from:faktury_prowizje@axaubezpieczenia.pl"

    for label in labels.items():
        results = service.users().messages().list(userId='me', labelIds=[label[1]], maxResults=2, q=query).execute()
        # Dwa różne maile z fv w tej samej labelce.
        # n = 1 if label[0] == 'Orange mob' and results['resultSizeEstimate'] > 1 else 0
        """Usł mobilne są w terminie pobrania drugim rezultatem, stąd n = 1. """
        n = 1 if label[0] == 'Orange mob' else 0
        message_id = ''
        try:
            message_id = results['messages'][n]['id']
        except Exception:
            pass
        msg = service.users().messages().get(userId='me', id=message_id).execute()

        yield label[0], message_id, msg


def attachment_id(fv, msg):
    """Sprawdza czy i o jakiej nazwie jest załącznik, przekazuje ID."""
    for part in msg['payload']['parts']:
        a = True if fv in ('AXA', 'Wiener', 'TUW', 'A-Z', 'AWS') and part['filename'] else False
        b = True if fv in ('Insly', 'Orange mob') and re.search('faktura', part['filename'], re.I) else False
        c = True if fv in ('Orange stac', 'TUZ') and re.search('.pdf$', part['filename'], re.I) else False
        if a or b or c and fv != 'Euroins':
            if 'data' in part['body']:
                att_id = part['body']['data']
                return att_id
            else:
                att_id = part['body']['attachmentId']
                return att_id


def attachment_id_gen(fv, msg):
    for part in msg['payload']['parts']:
        d = True if fv == 'Euroins' and re.search('(.pdf$|.zip)', part['filename']) else False
        e = True if fv == 'TUW' and re.search('(.pdf$|.zip)', part['filename']) else False
        if d or e:
            att_id = part['body']['attachmentId']
            yield att_id, part['filename']


def axa_invoice(fv, message_id, msg):
    if fv == 'AXA':
        if str(msg).find('plik prowizyjny') > -1:
            att_id = attachment_id(fv, msg)
            get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                   id=att_id).execute()
            get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
            path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/AXA_prowizja' + '.xls'])
            with open(path, 'wb') as f:
                f.write(get_att_de)

            # Ten fragment zdejmuje hasło z rozliczenia prowizyjnego AXA
            xlApp = Dispatch("Excel.Application")
            xlwb = xlApp.Workbooks.Open('C:\\Users\ROBERT\Desktop\Księgowość\\2021\RobO\AXA_prowizja.xls',
                                        False, False, None, 'PVxCC32%pLkO')
            path = ''.join(['C:\\Users\ROBERT\Desktop\Księgowość\\2021\RobO'])
            xlApp.DisplayAlerts = False
            xlwb.SaveAs(path + '\AXA_prowizja.xls', FileFormat=-4143, Password='')
            xlApp.DisplayAlerts = True
            xlwb.Close()
            print('AXA ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak AXA\n')
            print('Brak AXA')


def wiener_invoice(fv, message_id, msg):
    if fv == 'Wiener':
        if str(msg).find('prowizji za miesiąc') > -1:
            att_id = attachment_id(fv, msg)
            get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                   id=att_id).execute()
            get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
            path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/Wiener_prowizja' + '.pdf'])
            with open(path, 'wb') as f:
                f.write(get_att_de)
            print('Wiener ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak Wiener\n')
            print('Brak Wiener')


def insly_invoice(fv, message_id, msg):
    if fv == 'Insly':
        if str(msg).find('Faktura') > -1:
            att_id = attachment_id(fv, msg)
            get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                   id=att_id).execute()
            get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
            path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/Insly_faktura' + '.pdf'])
            with open(path, 'wb') as f:
                f.write(get_att_de)
            print('Insly ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak Insly\n')
            print('Brak faktury Insly')


def orange_stac_invoice(fv, message_id, msg):
    if fv == 'Orange stac':
        if '_' in str(msg['payload']['parts'][1]['filename']):
            att_id = attachment_id(fv, msg)
            get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                   id=att_id).execute()
            get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
            path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/Orange_faktura_stacjonarne' + '.pdf'])
            with open(path, 'wb') as f:
                f.write(get_att_de)
            print('Orange stacjonarne ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak Orange usługi stacjonarne\n')
            print('Brak faktury Orange usługi stacjonarne')


def orange_mobil_invoice(fv, message_id, msg):
    if fv == 'Orange mob':
        if 'FAKTURA' in str(msg['payload']['parts'][1]['filename']):
            att_id = attachment_id(fv, msg)
            get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                   id=att_id).execute()
            get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
            path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/Orange_faktura_mobilne' + '.pdf'])
            with open(path, 'wb') as f:
                f.write(get_att_de)
            print('Orange mobilne ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak Orange usługi mobilne\n')
            print('Brak faktury Orange usługi mobilne')


def aws_invoice(fv, message_id, msg):
    if fv == 'AWS':
        if str(msg).find('Invoice(s) available') > -1:
            att_id = attachment_id(fv, msg)
            get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                   id=att_id).execute()
            get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
            path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/AWS_faktura' + '.pdf'])
            with open(path, 'wb') as f:
                f.write(get_att_de)
            print('AWS ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak AWS\n')
            print('Brak faktury Amazon Web Services')


def tuw_invoice(fv, message_id, msg):
    if fv == 'TUW':
        """Raz wpisuje hasło w treść, raz nie. Powinien rozpoznawać pdf lub zip."""
        h = ''
        if re.search('Towarzystwo', str(msg)) or (h := re.search('hasło:\s?([A-z0-9!-_]+)', str(msg))):
            # att_id = attachment_id(fv, msg)
            for att_id, filename in attachment_id_gen(fv, msg):
                get_att = service.users().messages().attachments().get(userId='me',
                                                                       messageId=message_id,
                                                                       id=att_id).execute()
                get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
                """Raz wpisuje hasło w treść, raz nie."""
                if h:
                    path = ''.join([f'C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/TUW_faktura_haslo_{h.group(1)}'])
                else:
                    path = ''.join([f'C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/TUW_{filename}'])
                with open(path, 'wb') as f:
                    f.write(get_att_de)
                    # zip_ref = zipfile.ZipFile(path + '.zip')
                    # zip_ref.extractall(pwd='TUW!_5121_TUW'.encode('ascii'))
                if path + '.pdf':
                    print('TUW ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak TUW\n')
            print('Brak TUW')


def tuz_invoice(fv, message_id, msg):
    if fv == 'TUZ':
        if str(msg).find('zestawienie prowizyjne') > -1:
            att_id = attachment_id(fv, msg)
            get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                   id=att_id).execute()
            get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
            path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/TUZ_faktura'])
            with open(path + '.pdf', 'wb') as f:
                f.write(get_att_de)
            print('TUZ ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak TUZ\n')
            print('Brak TUZ')


def az_invoice(fv, message_id, msg):
    if fv == 'A-Z':
        if str(msg).find('fakturę') > -1:
            att_id = attachment_id(fv, msg)
            get_att = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                   id=att_id).execute()
            get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
            path = ''.join(['C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/AZ_faktura_haslo_Rozliczenia'])
            with open(path + '.zip', 'wb') as f:
                f.write(get_att_de)
            print('A-Z ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak A-Z\n')
            print('Brak A-Z')


def eins(fv, message_id, msg):
    if fv == 'Euroins':
        # print(msg)
        if re.search('(not[a|ę]+|prowizyjn[a|ą|y]+)', str(msg)):  # or 'Łuczak' in str(msg):
            for att_id, filename in attachment_id_gen(fv, msg):
                get_att = service.users().messages().attachments().get(userId='me',
                                                                       messageId=message_id,
                                                                       id=att_id).execute()
                get_att_de = base64.urlsafe_b64decode(get_att['data'].encode('UTF-8'))  # binary
                path = ''.join([f'C:/Users/ROBERT/Desktop/Księgowość/2021/RobO/Euroins_{filename}'.replace(' ', '_')])
                with open(path, 'wb') as f:
                    f.write(get_att_de)
                print('Euroins ok')
        else:
            with open(r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO\brak dokumentów.txt', 'a') as f:
                f.write('Brak Euroins\n')
            print('Brak noty Euroins')


def email():
    for fv, id, message in labels(service):
        # axa_invoice(fv, id, message)
        # wiener_invoice(fv, id, message)
        # insly_invoice(fv, id, message)
        # orange_stac_invoice(fv, id, message)
        # orange_mobil_invoice(fv, id, message)
        # aws_invoice(fv, id, message)
        tuw_invoice(fv, id, message)
        # tuz_invoice(fv, id, message)
        # az_invoice(fv, id, message)
        # eins(fv, id, message)


service = main()

if __name__ == '__main__':
    email()




