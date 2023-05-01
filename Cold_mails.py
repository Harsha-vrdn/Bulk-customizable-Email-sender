import smtplib
import ssl
import pandas as pd
from email.message import EmailMessage
from mimetypes import guess_type


def read_excel_file(file_path):
    """
    Reads the excel file and returns a pandas DataFrame object
    containing email addresses and names.
    Assumes that the first column contains email addresses and the second
    column contains names.
    """
    df = pd.read_excel(file_path, names=['MAIL', 'NAME'])
    return df


def get_email_and_name_lists(df):
    """
    Extracts email addresses and names from a pandas DataFrame object.
    Returns two lists: one containing email addresses and the other containing
    first names.
    """
    emails = df['MAIL'].values
    names = df['NAME'].values
    first_names = [name.split()[0] for name in names]
    return emails, first_names


def send_email(email_sender, email_password, recipient, subject, body, attachment_path=None):
    """
    Sends an email with or without attachment to the specified recipient.
    """
    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = recipient
    em['Subject'] = subject
    em.set_content(body)

    if attachment_path:
        file_name = attachment_path.split('/')[-1]
        mime_type, encoding = guess_type(file_name)
        app_type, sub_type = mime_type.split("/")[0], mime_type.split("/")[1]
        with open(attachment_path, 'rb') as f:
            file_data = f.read()
        em.add_attachment(file_data, maintype=app_type, subtype=sub_type, filename=file_name)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, recipient, em.as_string())
        print(f'Email sent to {recipient}')


if __name__ == '__main__':
    # Replace the file path, email_sender, and email_password with your own values.
    file_path = 'Your excel sheet name'
    email_sender = 'Your Email ID'
    email_password = 'Your 16 character app password'

    # Read the excel file and extract email addresses and names.
    df = read_excel_file(file_path)
    emails, first_names = get_email_and_name_lists(df)

    # Send an email to each recipient.
    subject = 'Regarding _______________'
    body = 'Hello {},\n\nHope you are doing Great.'
    attachment_path = 'Your attachment name'

    for recipient, name in zip(emails, first_names):
        body_formatted = body.format(name)
        send_email(email_sender, email_password, recipient, subject, body_formatted, attachment_path)

