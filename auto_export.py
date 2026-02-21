import imapclient
import pyzmail
import re
import requests
import os
import datetime
import smtplib
from email.message import EmailMessage

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ['https://www.googleapis.com/auth/drive']

EMAIL = "rutvik.Desai@smytten.com"
APP_PASSWORD = "lvxqnkaeewugzljb"

FORWARD_TO = [
    "",
    "rutvik.desai@smytten.com",
    "promotions@smytten.com",
    ""
]

folder_name = "Shop_Exports"
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# ===== DRIVE AUTH =====
creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
else:
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

drive_service = build('drive', 'v3', credentials=creds)

# ===== GMAIL READ =====
mail = imapclient.IMAPClient('imap.gmail.com', ssl=True)
mail.login(EMAIL, APP_PASSWORD)
mail.select_folder('INBOX')

messages = mail.search(['UNSEEN', 'SUBJECT', 'Shop Product Full Export'])

for uid in messages:
    raw_message = mail.fetch(uid, ['BODY[]'])
    message = pyzmail.PyzMessage.factory(raw_message[uid][b'BODY[]'])

    body = message.text_part.get_payload().decode(message.text_part.charset)

    match = re.search(r'https://[^\s"]+', body)
    if match:
        file_url = match.group(0)

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        readable_date = datetime.datetime.now().strftime("%d-%b-%Y")
        file_name = f"Shop_Product_Export_{timestamp}.xlsx"
        file_path = os.path.join(folder_name, file_name)

        print("Downloading file...")
        response = requests.get(file_url, stream=True)

        if response.status_code == 200:
            with open(file_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            print("Uploading to Drive...")
            file_metadata = {'name': file_name}
            media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

            file_id = uploaded_file.get('id')

            drive_service.permissions().create(
                fileId=file_id,
                body={'type': 'anyone', 'role': 'reader'}
            ).execute()

            drive_link = f"https://drive.google.com/file/d/{file_id}/view"

            # Send HTML Email
            msg = EmailMessage()
            msg['Subject'] = f"Shop Product Full Export | {readable_date}"
            msg['From'] = EMAIL
            msg['To'] = ", ".join(FORWARD_TO)

            html_content = f"""
<html>
<body style="margin:0; padding:0; font-family:sans-serif; color:#333; line-height:1.4;">
  <div style="font-size:13px; font-family:sans-serif;">
    <p style="margin:0;">Hello Everyone,</p>
    <p style="margin:8px 0;">PFA the today's <b>Shop Product Full Export</b>.</p>
    <p style="margin:8px 0;">
      <a href="{drive_link}" style="background-color:#004aad; color:white; text-decoration:none; padding:8px 14px; border-radius:6px; font-weight:bold; font-family:sans-serif;">
        🔗 Download File
      </a>
    </p>
    <p style="margin:14px 0 0 0;">&nbsp;</p> <!-- one blank line -->
    <div style="font-family:'Courier New', monospace; font-size:13px;">
      <p style="margin:0;">Thanks,</p>
      <p style="margin:0;"><b>Rutvik Desai</b></p>
    </div>
  </div>
</body>
</html>
"""

            msg.add_alternative(html_content, subtype='html')

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL, APP_PASSWORD)
                smtp.send_message(msg)

            print("HTML Email sent successfully!")

    mail.add_flags(uid, ['\\Seen'])

mail.logout()
print("All tasks completed successfully.")
