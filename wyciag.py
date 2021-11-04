
from __future__ import print_function
from ahk import AHK
import pyautogui
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
    time.sleep(9)
    for window in ahk.windows():
        if 'Santander' in window.title.decode('windows-1252'):
            win = window
            win.maximize()
            window.always_on_top = True
            return win


def log_into_account():
    # Tyko "Narzędzie Wycinanie" z Win10, screenshot z przeglądarki nie zadziała!
    path = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\wpłaty\images\\'
    if pyautogui.locateOnScreen(path + r'zalog_b.png'):
        pyautogui.click(path + r'zalog_b.png')
    else:
        pyautogui.locateOnScreen(path + r'zalog_sz.png')
        pyautogui.click(path + r'zalog_sz.png')

    pyautogui.locateOnScreen(path + r'sant_int.png')
    pyautogui.click(path + r'sant_int.png')

    time.sleep(5)
    ahk.run_script(f'Send, {numpad}')
    ahk.mouse_move(1297, 428, speed=10)
    ahk.click()
    time.sleep(3)
    ahk.click()
    time.sleep(8)


def download_summary(win):
    log_into_account()
    ahk.mouse_move(1676, 500, speed=10)
    ahk.click()
    time.sleep(6)
    ahk.mouse_wheel('down')
    ahk.mouse_wheel('down')
    ahk.mouse_move(1309, 970, speed=10)
    ahk.click()
    time.sleep(1)
    ahk.mouse_move(1023, 290, speed=10)
    ahk.click()
    time.sleep(1)
    ahk.mouse_move(1309, 970, speed=10)
    ahk.click()
    time.sleep(1)
    time.sleep(6)
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
    email_to = 'dg.jn@poczta.fm'
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
    project_path = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\księgowość\skrypty osobno\dist\\'
    if os.path.exists(project_path + 'token.pickle'):
        with open(project_path + 'token.pickle', 'rb') as token:
            creds = pickle.load(token)
    service = build('gmail', 'v1', credentials=creds)
    mail(service)
