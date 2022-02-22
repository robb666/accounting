
import os
import smtplib, ssl
from email import encoders
import mimetypes
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from dateutil.relativedelta import relativedelta
from L_H_ks import gapi


def send_attachments(sender_email, receiver_email):
    msc_rok = (datetime.today() + relativedelta(months=-1)).strftime('%m.%Y')
    message = MIMEMultipart()
    message['Subject'] = f'Test za {msc_rok}'
    body = """Cześć, to test smtplib.\n\n"""
    message.attach(MIMEText(body))

    documents = r'C:\Users\PipBoy3000\Desktop\Księgowość\Kadrowe\test'
    os.chdir(documents)
    for attachment in os.listdir(documents):
        content_type, encoding = mimetypes.guess_type(attachment, strict=False)
        print(content_type, encoding)
        if content_type is not None:
            main_type, sub_type = content_type.split('/', 1)
            my_file = MIMEBase(main_type, sub_type)
        else:
            pass

        with open(attachment, 'rb') as f:
            my_file.set_payload(f.read())
            my_file.add_header('Content-Disposition', f'attachment; filename = {attachment}',)
            encoders.encode_base64(my_file)
            message.attach(my_file)
            text = message.as_string()

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
        server.login('ubezpieczenia.magro@gmail.com', gapi)
        server.sendmail(sender_email, receiver_email, text)


send_attachments('ubezpieczenia.magro@gmail.com',
                 'robert.patryk.grzelak@gmail.com')