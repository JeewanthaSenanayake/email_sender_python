import pandas as pd
import time

from sender import EmailSender
from emailBodyAndDataList import getDataList

# # Set up the SMTP server details
# smtp_server = 'smtp.gmail.com' # This is for Gmail, When you use another mail server you should add their smtp_server
# smtp_port = 587 # This is for Gmail, When you use another mail server you should add their smtp_port
# smtp_username = 'ssbjms123@gmail.com' # Your email
# smtp_password = 'upsz pudp esqx tnqi' #enable two-step verification and search app password in google
# sender_name = 'Jeewantha Senanayake'

# # Read the Excel file into a DataFrame
# df = pd.read_excel('emails.xlsx')


# email_send = EmailSender()
# email_send.set_server(smtp_server=smtp_server,smtp_port=smtp_port,smtp_username=smtp_username,smtp_password=smtp_password,sender_name=sender_name)
# is_start = email_send.start_server()

# # Send the emails
# masg = """<h3>Hi {Name},</h3> <center><hr><h2>This is test email</h2><hr></center>"""
# masg+="""\nThank you."""
# dataList = getDataList(masg,df)

# if is_start:
#     for data in dataList:
#         email_send.send_email(data['email'], 'Test Email', data['body'])
#         # Add a 1-second delay
#         time.sleep(2)


import UI.ui as ui

app = ui.Application()
app.mainloop()