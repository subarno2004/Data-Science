#!/usr/bin/env python
# coding: utf-8

# In[1]:


pip install pyunpack patool


# In[2]:


import logging
import imaplib
import email
import ssl
import os
from pyunpack import Archive

SERVER = "imap.gmail.com"
USER = "Your_gmail_id"
APP_PASSWORD = "your_app_password" ## you need to enable 2-factor auth and create an app password
BATCH_SIZE = 10  # use a batch size for smooth handling

# Setup logging
logging.basicConfig(level=logging.DEBUG)

# Connect to the IMAP server
logging.debug('Connecting to ' + SERVER)
context = ssl.create_default_context()
server = imaplib.IMAP4_SSL(SERVER, ssl_context=context)

# Increase the timeout for reading email data
server.timeout = 120  # Set the timeout so that the app doesn't times out

# Login using the App Password
logging.debug('Logging in')
server.login(USER, APP_PASSWORD)

# Select the mailbox (e.g., "INBOX")
mailbox_name = "INBOX"
logging.debug(f'Selecting mailbox: {mailbox_name}')
server.select(mailbox_name)

# Search for all emails in the mailbox
logging.debug('Searching for emails')
status, email_ids = server.search(None, 'ALL')

# Create a directory to save attachments
attachment_dir = 'attachments'
if not os.path.exists(attachment_dir):
    os.makedirs(attachment_dir)

# Create a directory to extract .rar file contents
extracted_dir = 'extracted_contents'
if not os.path.exists(extracted_dir):
    os.makedirs(extracted_dir)

# Process emails in batches
email_id_list = email_ids[0].split()
for batch_start in range(0, len(email_id_list), BATCH_SIZE):
    batch_end = min(batch_start + BATCH_SIZE, len(email_id_list))
    batch_ids = email_id_list[batch_start:batch_end]

    for email_id in batch_ids:
        # Fetch the email
        logging.debug(f'Retrieving email with ID: {email_id}')
        status, email_data = server.fetch(email_id, '(RFC822)')

        # Parse the email content
        raw_email = email_data[0][1]
        message = email.message_from_bytes(raw_email)

        # Output message details
        print(f"From: {message['From']}")
        print(f"Subject: {message['Subject']}")
        print(f"Date: {message['Date']}")

        # Get the email body content
        email_body = ""
        for part in message.walk():
            if part.get_content_type() == "text/plain":
                email_body = part.get_payload(decode=True).decode('utf-8')
                break

        # Print or process the email body as needed
        print(f"Email Body:\n{email_body}\n")

        # Handle attachments of all types
        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                continue  # Skip multipart containers
            filename = part.get_filename()
            if filename:
                logging.debug(f"Saving attachment: {filename}")
                attachment_path = os.path.join(attachment_dir, filename)

                with open(attachment_path, 'wb') as attachment_file:
                    attachment_file.write(part.get_payload(decode=True))

                # Check if the attachment is a .rar file and extract its contents
                if filename.lower().endswith('.rar'):
                    logging.debug(f"Extracting .rar attachment: {filename}")
                    Archive(attachment_path).extractall(extracted_dir)

# Close the connection
server.logout()


# In[ ]:




