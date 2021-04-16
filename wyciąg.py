
from __future__ import print_function
from ahk import AHK
import time
import os
import pickle
import os.path
from googleapiclient.discovery import build
import base64
import mimetypes
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
from datetime import datetime
from dateutil.relativedelta import relativedelta
from L_H_ks import numpad

ahk = AHK()


def open_browser():
    ahk.run_script('Run, firefox.exe -new-window https://www.santander.pl/klient-indywidualny')
    time.sleep(4.5)
    for window in ahk.windows():
        if 'Santander' in window.title.decode('windows-1252'):
            win = window
            win.maximize()
            window.always_on_top = True
            return win


def log_into_account():
    ahk.mouse_move(1429, 157, speed=2)
    ahk.click()
    ahk.mouse_move(1404, 240, speed=2)
    ahk.click()
    time.sleep(3)
    ahk.run_script(f'Send, {numpad}')
    ahk.mouse_move(991, 402, speed=2)
    ahk.click()
    time.sleep(2)
    ahk.click()
    time.sleep(5)


def download_summary(win):
    log_into_account()
    ahk.mouse_move(1254, 471, speed=1)
    ahk.click()
    time.sleep(2)
    ahk.mouse_wheel('down')
    ahk.mouse_wheel('down')
    ahk.mouse_move(981, 939, speed=1)
    ahk.click()
    ahk.mouse_move(704, 260, speed=2)
    ahk.click()
    ahk.mouse_move(981, 939, speed=3)
    ahk.click()
    time.sleep(2)
    win.close()


def create_message(to, sender, subject, msg_text, msg_attachments):
    message = MIMEMultipart()

    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    part1 = MIMEText(msg_text, 'plain')
    message.attach(part1)

    for attachment in msg_attachments:
        file_name = os.path.basename(attachment)
        if file_name.endswith('mt940x'):
            my_file = MIMEBase(attachment, 'mt940x')
        else:
            content_type, encoding = mimetypes.guess_type(attachment, strict=False)
            main_type, sub_type = content_type.split('/', 1)
            file_name = os.path.basename(attachment)
            my_file = MIMEBase(main_type, sub_type)

        with open(attachment, 'rb') as f:
            my_file.set_payload(f.read())
            my_file.add_header('Content-Disposition', 'attachment', filename=file_name)
            encoders.encode_base64(my_file)
            message.attach(my_file)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return raw_message


def send_message(service, raw_message):
    try:
        message = service.users().messages().send(userId='me',
                                                  body={'raw': raw_message}).execute()
        return message
    except Exception as e:
        print('An error occurred: %s' % e)
        return None


def mail(service):
    """Email z wyciągiem."""
    msc_rok = (datetime.today() + relativedelta(months=-1)).strftime('%m.%Y')
    email_to = 'ubezpieczenia.magro@gmail.com'
    my_email = 'ubezpieczenia.magro@gmail.com'
    message_text = """
Cześć Ola,\n 
przesyłam wyciąg w formatach .mt940x i .pdf.\n\n\n
Pozdrawiam,
Robert Grzelak
tel.: 42 637 19 97
tel.kom.: 572 810 576\n
MAGRO UBEZPIECZENIA Sp. z o.o.
Spółka zarejestrowana pod numerem KRS 0000648004,
XX WYDZIAŁU GOSPODARCZEGO KRAJOWEGO REJESTRU SĄDOWEGO w Łodzi.
NIP 7252160008
www.ubezpieczenia-magro.pl
"""

    desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    attachments = []
    for item in os.listdir(desktop):
        if 'wyciag_' in item and os.path.isfile(os.path.join(desktop, item)):

            attachments.append(os.path.join(desktop, item))

    message = create_message(email_to, my_email, f'Wyciąg za {msc_rok}', message_text, attachments)
    send_message(service, message)

    [os.remove(os.path.join(desktop, file)) for file in attachments]


if __name__ == '__main__':
    win = open_browser()
    download_summary(win)
    # Program dodany do harmonogramu zadań wykonuje się z folderu C:\WINDOWS\system32 #
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    service = build('gmail', 'v1', credentials=creds)
    mail(service)




