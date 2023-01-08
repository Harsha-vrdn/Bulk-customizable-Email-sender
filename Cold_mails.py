import smtplib
import ssl
import pandas as pd
from email.message import EmailMessage
from mimetypes import guess_type

e = pd.read_excel('Your excel sheet name')

# make sure the rows are named MAIL and NAME.
emails = e['MAIL'].values
names = e['NAME'].values
newnameslist = []

# Use this for loop only if you want the first name to-be displayed in the Email.
for fullname in names:
    first=fullname.split()[0]
    last=fullname.split()[-1]
    tempname = first
    newnameslist.append(tempname)


combinedlist = list(zip(newnameslist,emails))
#This is only if you want to add an attachment
file_name = "Your attachment name"
mime_type, encoding = guess_type(file_name)
app_type, sub_type = mime_type.split("/")[0], mime_type.split("/")[1]

for (i,n) in combinedlist:
    # Make sure you have your 2 step verification on your gmail account
    # Make sure you have a 16 character app password
    email_sender = 'Your Email ID'
    email_password = 'Your 16 character app password'


    subject = 'Regarding _______________'
    body = f'Hello {i}, \n\nHope you are doing Great.'

    em = EmailMessage()
    em['From']= email_sender
    em['To']= n
    em['Subject']= subject
    em.set_content(body)
    with open('Your attachment name', 'rb') as f:
        file_data = f.read()
    em.add_attachment(file_data, maintype=app_type, subtype=sub_type, filename=file_name)
    

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context= context) as smtp:
        smtp.login(email_sender,email_password)
        smtp.sendmail(email_sender, n, em.as_string())
        print(f'Email sent to {n}')
