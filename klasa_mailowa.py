
import os
import email, smtplib, ssl
from email import encoders
import mimetypes
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64
from L_H_ks import gapi

subject = 'Dokumenty za xxx'
body = 'Cześć, przesyłam dokumenty w załącznikach.'
smtp_server = 'smtp.gmail.com'
sender_email = 'ubezpieczenia.magro@gmail.com'
receiver_email = 'ubezpieczenia.magro@gmail.com'

# message = MIMEMultipart('alternative')
message = MIMEMultipart()
message['From'] = sender_email
message['To'] = receiver_email
message['Subject'] = 'Teściowa'

# Add body to email
message.attach(MIMEText(body, 'plain'))

documents = r'C:\Users\ROBERT\Desktop\Księgowość\2021\RobO'
# filename = "document.txt"  # In same directory as script
attachments = []

for attachment in os.listdir(documents):
    print(attachment)
    attachments.append(os.path.join(documents, attachment))

for attachment in attachments:
    file_name = os.path.basename(attachment)
    print(documents, file_name)
    content_type, encoding = mimetypes.guess_type(attachment, strict=False)
    main_type, sub_type = content_type.split('/', 1)

    my_file = MIMEBase(main_type, sub_type)



    with open(attachment, 'rb') as f:
        part = MIMEBase('application', 'octet-stream', filename=file_name)
        my_file.set_payload(f.read())




        part.add_header('Content-Disposition',
                        f'attachment; filename = {file_name}',)

        encoders.encode_base64(my_file)

        message.attach(my_file)
        text = message.as_string()



context = ssl.create_default_context()
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
    server.login('ubezpieczenia.magro@gmail.com', gapi)
    server.sendmail(sender_email, receiver_email, text)



