import sys
import imaplib
import email
from email.header import decode_header
from bs4 import BeautifulSoup
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QCheckBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtCore import QUrl

class EmailUnsubscriber(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set main window properties
        self.setWindowTitle("Unsubem")
        self.setGeometry(100, 100, 600, 400)

        # Create central widget and layout
        self.central_widget = QWidget(self)
        self.layout = QVBoxLayout()

        # Add dropdown for email providers
        self.provider_label = QLabel("Email Provider:", self)
        self.layout.addWidget(self.provider_label)
        self.provider_dropdown = QComboBox(self)
        self.provider_dropdown.addItems(["Gmail", "Outlook"])
        self.layout.addWidget(self.provider_dropdown)

        # Add input fields for email and password
        self.email_label = QLabel("Email:", self)
        self.layout.addWidget(self.email_label)
        self.email_input = QLineEdit(self)
        self.layout.addWidget(self.email_input)

        self.password_label = QLabel("Password:", self)
        self.layout.addWidget(self.password_label)
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.layout.addWidget(self.password_input)

        # Create table for displaying emails
        self.email_table = QTableWidget(self)
        self.email_table.setColumnCount(4)
        self.email_table.setHorizontalHeaderLabels(["Select", "Email Subject", "Unsubscribe Link", "Status"])
        header = self.email_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)  # Make headers resizable
        self.layout.addWidget(self.email_table)

        # Button to start fetching emails
        self.fetch_button = QPushButton("Fetch Unsubscribe Links", self)
        self.fetch_button.clicked.connect(self.fetch_emails)
        self.layout.addWidget(self.fetch_button)

        # Button to open selected unsubscribe links
        self.unsubscribe_button = QPushButton("Open Selected Unsubscribe Links", self)
        self.unsubscribe_button.clicked.connect(self.open_unsubscribe_links)
        self.layout.addWidget(self.unsubscribe_button)

        # Status label
        self.status_label = QLabel("Overall Status:", self)
        self.layout.addWidget(self.status_label)

        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

    def fetch_emails(self):
        # Get the selected email provider
        selected_provider = self.provider_dropdown.currentText()

        # Define IMAP settings for Gmail and Outlook
        imap_settings = {
            "Gmail": "imap.gmail.com",
            "Outlook": "imap-mail.outlook.com"
        }

        # Get the IMAP server for the selected provider
        imap_server = imap_settings.get(selected_provider, "imap.gmail.com")  # Default to Gmail if not found

        # Logging into the email using imaplib
        email_address = self.email_input.text()
        email_password = self.password_input.text()

        # Create an IMAP4 SSL session
        mail = imaplib.IMAP4_SSL(imap_server)
        try:
            # Login to the account
            mail.login(email_address, email_password)

            # Select the mailbox you want to check (INBOX in this case)
            mail.select("inbox")

            # Search for all emails
            status, message_numbers = mail.search(None, "ALL")

            # Get all email IDs
            email_ids = message_numbers[0].split()

            # Keywords to check in email body in multiple languages
            unsubscribe_keywords = ["unsubscribe", "désabonnement", "uitschrijven"]

            # Clear the table before adding new rows
            self.email_table.setRowCount(0)

            # Loop through each email ID to fetch the email details
            for e_id in email_ids:
                _, msg_data = mail.fetch(e_id, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject, encoding = decode_header(msg["Subject"])[0]
                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding if encoding else 'utf8')
                        body = ""
                        # Check if email is multipart
                        if msg.is_multipart():
                            # Iterate over email parts
                            for part in msg.walk():
                                content_type = part.get_content_type()
                                content_disposition = str(part.get("Content-Disposition"))
                                # Extract only the text parts
                                if "attachment" not in content_disposition:
                                    body += part.get_payload(decode=True).decode()
                        else:
                            body = msg.get_payload(decode=True).decode()

                        # Check if any of the unsubscribe keywords are in the email body
                        if any(keyword in body.lower() for keyword in unsubscribe_keywords):
                            # Extract the unsubscribe link (simplified for demonstration)
                            soup = BeautifulSoup(body, "html.parser")
                            # Find the link that contains any of the unsubscribe keywords
                            unsubscribe_link = next((a["href"] for a in soup.find_all("a") if any(keyword in a.get_text().lower() for keyword in unsubscribe_keywords)), None)
                            if unsubscribe_link:
                                # Add the subject and unsubscribe link to the table
                                self.add_to_table(subject, unsubscribe_link)

            # Update the status label
            self.status_label.setText(f"Overall Status: Fetched {self.email_table.rowCount()} unsubscribe links")

            # Logout and close the connection
            mail.logout()
        except Exception as e:
            # Display error if any (for debugging purposes)
            print(e)
            self.status_label.setText(f"Overall Status: Error - {str(e)}")

    def add_to_table(self, subject, unsubscribe_link):
        # Insert a new row at the end of the table
        row_count = self.email_table.rowCount()
        self.email_table.insertRow(row_count)

        # Create a checkbox for selecting the row
        checkbox = QCheckBox()
        checkbox.setCheckState(Qt.Checked)  # Set the checkbox to be checked by default
        self.email_table.setCellWidget(row_count, 0, checkbox)

        # Add the subject and unsubscribe link to the new row
        self.email_table.setItem(row_count, 1, QTableWidgetItem(subject))
        self.email_table.setItem(row_count, 2, QTableWidgetItem(unsubscribe_link))
        self.email_table.setItem(row_count, 3, QTableWidgetItem("Not Opened"))  # Set initial status

    def open_unsubscribe_links(self):
        # Iterate over each row in the table
        for row in range(self.email_table.rowCount()):
            # Check if the checkbox is checked
            if self.email_table.cellWidget(row, 0).checkState() == Qt.Checked:
                # Get the unsubscribe link from the table
                unsubscribe_link = self.email_table.item(row, 2).text()
                
                try:
                    # Open the unsubscribe link in the default web browser
                    QDesktopServices.openUrl(QUrl(unsubscribe_link))
                    # Update the status in the table
                    self.email_table.setItem(row, 3, QTableWidgetItem("Opened"))
                except Exception as e:
                    # Update the status in the table if an error occurs
                    self.email_table.setItem(row, 3, QTableWidgetItem(f"Error: {str(e)}"))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = EmailUnsubscriber()
    main_window.show()
    sys.exit(app.exec_())