import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import Menu
import os
import pandas as pd
import time
from sender import EmailSender
from emailBodyAndDataList import getDataList

# Function to handle focus in event


def on_focus_in(entry, placeholder):
    if entry.get() == placeholder:
        entry.delete(0, tk.END)
        entry.config(fg='black')

# Function to handle focus out event


def on_focus_out(entry, placeholder):
    if not entry.get():
        entry.insert(0, placeholder)
        entry.config(fg='grey')


def on_focus_in_text(text_widget, placeholder):
    if text_widget.get("1.0", "end-1c") == placeholder:
        text_widget.delete("1.0", tk.END)
        text_widget.config(fg='black')


def on_focus_out_text(text_widget, placeholder):
    if not text_widget.get("1.0", "end-1c").strip():
        text_widget.insert("1.0", placeholder)
        text_widget.config(fg='grey')


class Application(tk.Tk):

    uploded_file_path = "None"
    # Set up the SMTP server details
    # This is for Gmail, When you use another mail server you should add their smtp_server
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587  # This is for Gmail, When you use another mail server you should add their smtp_port
    smtp_username = ''  # Your email
    smtp_password = ''  # enable two-step verification and search app password in google
    sender_name = ''

    def __init__(self):
        super().__init__()

        self.title("E-mail Sender")

        # Set the window to full screen and disable resizing
        # Set a fixed window size
        self.geometry("720x620")  # Width x Height
        self.resizable(False, False)

        # Set background color to black
        self.configure(bg='black')

        # Create frames
        self.login_frame = tk.Frame(self, bg='black')
        self.greeting_frame = tk.Frame(self, bg='black')

        # Check if the file exists
        if os.path.exists("database\email_data.txt"):
            # Initialize login frame
            self.create_greeting_frame()
            self.greeting_frame.pack(fill="both", expand=True)

            with open('database/email_data.txt', 'r') as file:
                # Read all lines into a list
                user_data = file.readlines()
                self.smtp_username = user_data[1]
                self.smtp_password = user_data[2]
                self.sender_name = user_data[0]
        else:
            # Initialize login frame
            self.create_login_frame()
            self.login_frame.pack(fill="both", expand=True)

        # start smtp server
        self.email_send = EmailSender()
        self.email_send.set_server(smtp_server=self.smtp_server, smtp_port=self.smtp_port,
                                   smtp_username=self.smtp_username, smtp_password=self.smtp_password, sender_name=self.sender_name)
        try:
            self.is_start = self.email_send.start_server()
            if (self.is_start == False):
                messagebox.showerror(
                    "Error", "Email and Password Mismatch or Connection issue")
                exit()
        except:
            messagebox.showerror(
                "Error", "Email and Password Mismatch or Connection issue")
            exit()

        menu = Menu(self)
        self.config(menu=menu)
        filemenu = Menu(menu)
        menu.add_cascade(label='File', menu=filemenu)
        filemenu.add_command(label='Logout', command=self.log_out)
        filemenu.add_separator()
        filemenu.add_command(label='Exit', command=self.quit)
        helpmenu = Menu(menu)
        menu.add_cascade(label='Help', menu=helpmenu)
        helpmenu.add_command(label='About')

    def create_login_frame(self):
        label_username = tk.Label(self.login_frame, text="Account Details", font=(
            "Helvetica", 30), bg='black', fg='white')
        label_username.pack(pady=20)

        # Username
        label_username = tk.Label(
            self.login_frame, text="Username", bg='black', fg='white')
        label_username.pack(pady=5)
        self.entry_username = tk.Entry(self.login_frame, width=40)
        self.entry_username.pack(pady=5)

        # Email
        label_email = tk.Label(
            self.login_frame, text="Email", bg='black', fg='white')
        label_email.pack(pady=5)
        self.entry_email = tk.Entry(self.login_frame, width=40)
        self.entry_email.pack(pady=5)

        # Password
        label_password = tk.Label(
            self.login_frame, text="Password", bg='black', fg='white')
        label_password.pack(pady=5)
        self.entry_password = tk.Entry(self.login_frame, width=40, show='*')
        self.entry_password.pack(pady=5)

        # Login Button
        login_button = tk.Button(
            self.login_frame, text="Login", width=20, command=self.login)
        login_button.pack(pady=25)

    def create_greeting_frame(self):
        label_username = tk.Label(
            self.greeting_frame, text="Write E-mail", font=("Helvetica", 30), bg='black', fg='white')
        label_username.pack(pady=15)

        # Create a button to trigger the file dialog
        file_button = tk.Button(
            self.greeting_frame, text="Select Excel File", command=self.open_file)
        file_button.pack(pady=5)

        # email subject
        placeholder_subject = "Subject"
        self.entry_subject = tk.Entry(self.greeting_frame, width=80)
        self.entry_subject.pack(pady=10)
        self.entry_subject.insert(0, placeholder_subject)

        self.entry_subject.bind("<FocusIn>", lambda event: on_focus_in(
            self.entry_subject, placeholder_subject))
        self.entry_subject.bind("<FocusOut>", lambda event: on_focus_out(
            self.entry_subject, placeholder_subject))

        # email body
        self.text_area_email_body = tk.Text(
            self.greeting_frame, width=80, height=20)
        self.text_area_email_body.pack(pady=5)

        placeholder_email_body = "Enter email"
        self.text_area_email_body.insert("1.0", placeholder_email_body)

        # Bind events for focus in and out
        self.text_area_email_body.bind("<FocusIn>", lambda event: on_focus_in_text(
            self.text_area_email_body, placeholder_email_body))
        self.text_area_email_body.bind("<FocusOut>", lambda event: on_focus_out_text(
            self.text_area_email_body, placeholder_email_body))

        # Create a button to trigger the file dialog
        file_button = tk.Button(self.greeting_frame,
                                text="Send E-mail", command=self.send_email)
        file_button.pack(pady=10)

        # lable
        self.send_status_label = tk.Label(self.greeting_frame, text="", font=(
            "Helvetica", 10), bg='black', fg='white')
        self.send_status_label.pack(pady=10)


# Functions

    def login(self):
        username = self.entry_username.get()
        email = self.entry_email.get()
        password = self.entry_password.get()
        # Open a file in write mode ('w'). This will create the file if it doesn't exist.
        with open('database/email_data.txt', 'w') as file:
            file.write(f"{username}\n")
            file.write(f"{email}\n")
            file.write(f"{password}\n")

        # Simple check to ensure fields are filled
        if not username or not email or not password:
            messagebox.showerror("Error", "All fields are required!")
        else:
            # Switch to the greeting frame
            self.login_frame.pack_forget()
            self.create_greeting_frame()
            self.greeting_frame.pack(fill="both", expand=True)

    def open_file(self):
        file_path = filedialog.askopenfilename(
            title="Select a Excel File",
            filetypes=(("Excel files", "*.xlsx*"), ("All files", "*.xlsx*"))
            # filetypes=(("Excel files", "*.xlsx*"))
        )

        # Display the file path or process the file
        if file_path:
            # print(f"Selected file: {file_path}")
            self.uploded_file_path = file_path
        else:
            messagebox.showerror("Error", "Exel is not uploaded!")


        
    def send_email(self):
        self.send_status_label.config(text="Sending...")
        if (self.uploded_file_path == "None" or self.text_area_email_body.get("1.0", "end-1c") == ""
           or self.text_area_email_body.get("1.0", "end-1c") == "Enter email"
           or self.entry_subject.get() == "" or self.entry_subject.get() == "Subject"):
            messagebox.showerror(
                "Error", "Select Exel and Type the Email Body!")
        else:
            e_body = self.text_area_email_body.get("1.0", "end-1c")
            e_subject = self.entry_subject.get().strip()

            df = pd.read_excel(self.uploded_file_path)
            masg = f"""{e_body}"""
            dataList = getDataList(masg, df)

            if self.is_start:
                for data in dataList:
                    e_status = self.email_send.send_email(
                        data['email'], e_subject, data['body'])
                    # Add a 1-second delay
                    self.send_status_label.config(text=e_status)
                    time.sleep(1)
            self.send_status_label.config(text="All e-mails are Sucsessfuly Sent")
                    
                  

    def log_out(self):
        file_name = "database\email_data.txt"
        # Delete the file
        if os.path.exists(file_name):
            os.remove(file_name)
            print(f"{file_name} has been deleted.")
            self.greeting_frame.pack_forget()
            self.login_frame.pack(fill="both", expand=True)
        else:
            messagebox.showerror("Error", "You are not loged!")
