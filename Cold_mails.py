import smtplib
import ssl
import pandas as pd
from email.message import EmailMessage
from mimetypes import guess_type


def send_mails(mylist):
    
    # Get the MIME type and encoding of the file to be attached
    file_name = "Your attachment name"
    mime_type, encoding = guess_type(file_name)
    app_type, sub_type = mime_type.split("/")[0], mime_type.split("/")[1]

    # Replace the following placeholders with your own email address and app password
    email_sender = 'Your Email ID'
    email_password = 'Your 16 character app password'

    # Loop through the list of names and emails and send emails to each recipient
    for (i, n) in mylist:
        # Replace the following placeholders with your own email address and app password
        email_sender = 'Your Email ID'
        email_password = 'Your 16 character app password'

        # Set the subject and body of the email
        subject = 'Regarding _______________'
        body = f'Hello {i}, \n\nHope you are doing great.'

        # Create a new EmailMessage object and attach the file to it
        em = EmailMessage()
        em['From'] = email_sender
        em['To'] = n
        em['Subject'] = subject
        em.set_content(body)
        with open(file_name, 'rb') as f:
            file_data = f.read()
        em.add_attachment(file_data, maintype=app_type, subtype=sub_type, filename=file_name)

        # Use SSL to connect to the SMTP server and send the email
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(email_sender, email_password)
            smtp.sendmail(email_sender, n, em.as_string())
            print(f'Email sent to {n}')


# Read the email addresses and names from an Excel sheet
e = pd.read_excel('Your excel sheet name')
emails = e['MAIL'].values
names = e['NAME'].values

# Use this for loop only if you want the first name to-be displayed in the email
newnameslist = []
for fullname in names:
    first = fullname.split()[0]
    last = fullname.split()[-1]
    tempname = first
    newnameslist.append(tempname)

# Combine the names and email addresses into a list
combinedlist = list(zip(newnameslist, emails))

# Call the send_mails function to send the emails
send_mails(combinedlist)
