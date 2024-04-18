import imaplib
import email
import re
import requests
from email.header import decode_header
import tkinter as tk
from tkinter import messagebox
import threading

class EmailUnsubscriber(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('Email Unsubscriber')

        # input fields for email provider, email, and password
        self.provider_label = tk.Label(self, text='Email Provider')
        self.provider_label.pack()

        self.provider_entry = tk.Entry(self)
        self.provider_entry.pack()

        self.email_label = tk.Label(self, text='Email')
        self.email_label.pack()

        self.email_entry = tk.Entry(self)
        self.email_entry.pack()

        self.password_label = tk.Label(self, text='Password')
        self.password_label.pack()

        self.password_entry = tk.Entry(self, show='*')
        self.password_entry.pack()

        # listbox to display emails
        self.email_listbox = tk.Listbox(self, selectmode=tk.MULTIPLE)
        self.email_listbox.pack()

        # button to start process
        self.start_button = tk.Button(self, text='Start', command=self.start)
        self.start_button.pack()

        # text area to display live report
        self.report_text = tk.Text(self)
        self.report_text.pack()

        # dict to store email ids and corresponding unsubscribe links
        self.emails = {}

    def start(self):
        # get email provider, email, and password
        provider = self.provider_entry.get()
        email = self.email_entry.get()
        password = self.password_entry.get()

        # check if the provider is supported
        if provider.lower() not in ['gmail', 'outlook']:
            messagebox.showerror('Error', 'Unsupported email provider')
            return

        # start unsubscribe process in separate thread
        threading.Thread(target=self.unsubscribe, args=(provider, email, password)).start()

    def unsubscribe(self, provider, user_email, password):
        # connect to email server
        if provider.lower() == 'gmail':
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
        elif provider.lower() == 'outlook':
            mail = imaplib.IMAP4_SSL("imap-mail.outlook.com")
        else:
            return

        try:
            mail.login(user_email, password)
        except imaplib.IMAP4.error:
            messagebox.showerror('Error', 'Login failed')
            return

        # select the mailbox you want to delete in
        box = 'inbox'
        try:
            mail.select(box)
        except imaplib.IMAP4.error:
            messagebox.showerror('Error', 'Failed to select mailbox')
            return

        # get uids
        resp, items = mail.uid('search', None, "ALL")
        if resp != 'OK':
            messagebox.showerror('Error', 'Failed to fetch emails')
            return

        email_ids = items[0].split()  # getting mails id

        for i in email_ids:
            # fetch the email body (RFC822) for the given ID
            result, data = mail.uid('fetch', i, '(BODY.PEEK[])')
            if result != 'OK':
                messagebox.showerror('Error', 'Failed to fetch email')
                return

            raw_email = data[0][1].decode("utf-8")  # decode mail
            email_message = email.message_from_string(raw_email)

           # ...

        # decode the email subject and add to listbox
        subject, encoding = decode_header(email_message['Subject'])[0]
        if isinstance(subject, bytes):
            if encoding is not None:
                subject = subject.decode(encoding)
            else:
                subject = subject.decode()

# ...


            # look for unsubscribe link in email body
            unsubscribe_link = None
            if email_message.is_multipart():
                for part in email_message.walk():
                    if part.get_content_type() == "text/html":
                        body = part.get_payload(decode=True)
                        links = re.findall(r'href=[\'"]?([^\'" >]+)', str(body))
                        for link in links:
                            if "unsubscribe" in link:
                                unsubscribe_link = link
                                break
            else:
                body = email_message.get_payload(decode=True)
                links = re.findall(r'href=[\'"]?([^\'" >]+)', str(body))
                for link in links:
                    if "unsubscribe" in link:
                        unsubscribe_link = link
                        break

            # store the email id and corresponding unsubscribe link
            self.emails[i] = unsubscribe_link

        # disconnect from the server
        mail.logout()

        # check which emails are selected in the listbox
        selected_indices = self.email_listbox.curselection()

        # unsubscribe from the selected emails
        for index in selected_indices:
            email_id = email_ids[index]
            unsubscribe_link = self.emails[email_id]

            if unsubscribe_link is not None:
                try:
                    response = requests.get(unsubscribe_link)
                    if response.status_code == 200:
                        print(f"Successfully unsubscribed: {unsubscribe_link}")
                    else:
                        print(f"Failed to unsubscribe: {unsubscribe_link}")
                except:
                    print(f"Error occurred while unsubscribing: {unsubscribe_link}")
            else:
                print(f"No unsubscribe link found for email {email_id}")

if __name__ == '__main__':
    app = EmailUnsubscriber()
    app.mainloop()
