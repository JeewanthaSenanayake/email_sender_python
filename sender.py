import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

class EmailSender:
    smtp_server = ''
    smtp_port = ''
    smtp_username = ''
    smtp_password = ''
    sender_name = ''

    EmailServer = None

    def set_server(self, smtp_server, smtp_port, smtp_username, smtp_password, sender_name):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.smtp_username = smtp_username
        self.smtp_password = smtp_password
        self.sender_name = sender_name

    def start_server(self):
        # Connect to the SMTP server
        try:
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.smtp_username, self.smtp_password)
            print('Connected to the SMTP server successfully!')
            self.EmailServer = server
            return True
        except smtplib.SMTPException as e:
            print('Error connecting to the SMTP server:', str(e))
            return False

    def send_email(self, recipient_email, subject, body,attachments_fils):

        # Add your name and email subject
        msg = MIMEMultipart()
        msg['From'] = self.sender_name
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # body = bodyCreate(recipient)
        msg.attach(MIMEText(body, 'html'))

        # Attach files
        if(len(attachments_fils)>0):
            for file_path in attachments_fils:
                # Open the file to be sent
                with open(file_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())

                # Encode the file in ASCII characters to send by email    
                encoders.encode_base64(part)

                # Add header as key/value pair to attachment part
                part.add_header('Content-Disposition', f'attachment; filename= {file_path.split("/")[-1]}')

                # Attach the part to the email message
                msg.attach(part)

        self.EmailServer.send_message(msg)
        
        print(f'Email sent successfully to {recipient_email}!')
        return f'Email sent successfully to {recipient_email}!'
        
