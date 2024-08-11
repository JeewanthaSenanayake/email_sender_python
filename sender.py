import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


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

    def send_email(self, recipient_email, subject, body):

        # Add your name and email subject
        msg = MIMEMultipart()
        msg['From'] = self.sender_name
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # body = bodyCreate(recipient)
        msg.attach(MIMEText(body, 'html'))

        self.EmailServer.send_message(msg)
        print(f'Email sent successfully to {recipient_email}!')
        return f'Email sent successfully to {recipient_email}!'
        
