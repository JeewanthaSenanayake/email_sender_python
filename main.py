import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import time

# Read the Excel file into a DataFrame
df = pd.read_excel('emails.xlsx')

dataList =[]

for index, row in df.iterrows():
    # values from three columns
    email = row['Email']
    name = row['Name']

    dataList.append( {'email': email, 'name': name})



# Set up the SMTP server details
smtp_server = 'smtp.gmail.com' # This is for Gmail, When you use another mail server you should add their smtp_server
smtp_port = 587 # This is for Gmail, When you use another mail server you should add their smtp_port
smtp_username = '***YOUR EMAIL***'
smtp_password = '***YOUR PASSWORD***' #enable two-step verification and get security key

def bodyCreate(data):
    # create your email body hear
    uname = data['name']
    
    masg = "<h3>Hi "+ uname +",</h3> <center><hr><h2>This is test email</h2><hr></center>"
    masg+="\nThank you."   
    return masg


# Connect to the SMTP server and send the email
try:
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        for recipient in dataList:
            
            recipient_email = recipient['email']

            # Add your name and email subject
            msg = MIMEMultipart()
            msg['From'] = '***YOUR NAME***'
            msg['To'] = recipient_email
            msg['Subject'] = '***SUBJECT OF THE EMAIL***'

        
            body = bodyCreate(recipient)
            msg.attach(MIMEText(body, 'html'))

            server.send_message(msg)
            print(f'Email sent successfully to {recipient_email}!')
            # Add a 1-second delay
            time.sleep(2)
except smtplib.SMTPException as e:
    print('Error sending email:', str(e))
