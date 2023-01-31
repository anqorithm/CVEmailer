from email.mime.text import MIMEText
import smtplib
from decouple import config
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd


def send_email(subject, sender_email, reciver_email, email_contnet, attachments_path_array, auth_email, auth_app_password):
    email_message = MIMEMultipart()
    email_message['Subject'] = subject
    email_message['From'] = sender_email
    email_message['To'] = reciver_email
    for attachment_path in attachments_path_array:
        attachment = MIMEBase('application', "octet-stream")
        attachment.set_payload(open(attachment_path, "rb").read())
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition',
                              'attachment', filename=get_file_original_name(attachment_path))
        email_message.attach(attachment)
    email_message.attach(MIMEText(email_contnet, 'plain'))
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(auth_email, auth_app_password)
    server.send_message(email_message)
    server.quit()


def get_file_original_name(attachment_path):
    return attachment_path.split('/')[1]


def main():
    auth_app_password = config('AUTH_APP_PASSWORD')
    auth_email = config('AUTH_EMAIL')
    sender_email = config('SENDER_EMAIL')
    attachments = ['attachments/dummy_cv.pdf', 'attachments/dummy_cv.txt']
    email_subject = 'COOP Training'
    email_content = '''I'm applying to your company because it is awesome 🤩'''
    df = pd.read_excel('emails.xlsx')
    for row in df.iterrows():
        reciver_email = row[1]['Email']
        send_email(email_subject, sender_email, reciver_email,
                   email_content, attachments, auth_email, auth_app_password)
        print(f'You have sent an email to {reciver_email} successfully')


if __name__ == "__main__":
    main()
