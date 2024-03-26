import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import csv
import re


# Set the maximum line length
imaplib._MAXLINE = 10000000  # Increase to 10 MB (adjust as needed)

# IMAP server settings
imap_server = 'imap.zoho.com'
username = 'konnect@ahadicorp.com'
password = 'K@nn3ctL@v3'

# Open the CSV file in append mode
with open('fetched_emails.csv', 'a', newline='') as csvfile:
    writer = csv.writer(csvfile)

    # Connect to the IMAP server
    imap = imaplib.IMAP4_SSL(imap_server)

    # Login to the server
    imap.login(username, password)

    # Select the mailbox (inbox)
    imap.select('inbox')

    # Calculate the time for the past 24 hours
    since_date = (datetime.now() - timedelta(days=1)).strftime('%d-%b-%Y')

    # Search for emails from the last 24 hours
    status, messages = imap.search(None, '(SINCE "{}")'.format(since_date))

    if status == 'OK':
        # Iterate through each email
        for message_id in messages[0].split():
            # Fetch the email
            status, msg_data = imap.fetch(message_id, '(RFC822)')
            if status == 'OK':
                # Parse the email
                raw_email = msg_data[0][1]
                email_message = email.message_from_bytes(raw_email)
                # Get email details
                subject = decode_header(email_message['Subject'])[0][0]
                sender = email.utils.parseaddr(email_message['From'])[1]
                date_sent = email.utils.parsedate_to_datetime(email_message['Date'])
                # Extract email body
                email_body = None
                for part in email_message.walk():
                    if part.get_content_type() == "text/html":
                        html_content = part.get_payload(decode=True).decode(part.get_content_charset())
                        # Use BeautifulSoup to extract text from HTML
                        soup = BeautifulSoup(html_content, 'html.parser')
                        email_body = soup.get_text(separator='\n').strip()  # Strip leading and trailing whitespace
                        # Remove excess whitespace
                        email_body = re.sub(r'\n\s*\n', '\n\n', email_body)
                        break

                # Write data to CSV file
                if email_body is not None:
                    writer.writerow([subject, sender, date_sent, email_body])
                else:
                    print("Warning: Email body is empty, skipping...")

                # Print email details (optional for verification)
                print('Subject:', subject)
                print('From:', sender)
                print('Date:', date_sent)
                print('Body:', email_body)
                print('-' * 50)
            else:
                print("Failed to fetch email with message ID:", message_id)

    # Close the connection
    imap.close()
    imap.logout()
