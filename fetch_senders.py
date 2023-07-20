import imaplib
import email
import csv
from datetime import datetime
from tqdm import tqdm  # Import tqdm for the progress bar
from email.header import decode_header
from email.utils import parseaddr
from config import IMAP_SERVER, USERNAME, PASSWORD, MAILBOX, start_date, end_date

# Connect to the IMAP server
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(USERNAME, PASSWORD)

# Select the mailbox
mail.select(MAILBOX)

# Convert dates to IMAP-compliant format (e.g., 01-Jun-2023)
start_date_str = (datetime.strptime(start_date, '%Y-%m-%d')).strftime('%d-%b-%Y')
end_date_str = (datetime.strptime(end_date, '%Y-%m-%d')).strftime('%d-%b-%Y')

# Custom search query for date range
search_query = f'(SENTSINCE {start_date_str} SENTBEFORE {end_date_str})'

# Search for emails in the date range
status, email_ids = mail.search(None, search_query)

sender_data = set()  # Use a set to automatically remove duplicates

# Function to properly decode the sender's name
def decode_sender_name(name):
    decoded_name = ""
    for part, encoding in decode_header(name):
        if isinstance(part, bytes):
            decoded_name += part.decode(encoding or "utf-8")
        else:
            decoded_name += part
    return decoded_name.strip()

# Fetch and parse emails with tqdm progress bar
total_emails = len(email_ids[0].split())
for email_id in tqdm(email_ids[0].split(), desc='Processing emails', unit='email', unit_scale=True):
    status, email_data = mail.fetch(email_id, '(RFC822)')
    raw_email = email_data[0][1]
    msg = email.message_from_bytes(raw_email)

    # Decode the sender's name
    sender_name = decode_sender_name(msg["From"])

    # Remove the sender email address from the first column
    sender_name = sender_name.split('<')[0].strip()

    # Extract the sender email address
    _, sender_email = parseaddr(msg["From"])

    # Convert email address to lowercase for case-insensitive uniqueness
    sender_email = sender_email.lower()

    sender_data.add((sender_name, sender_email))

# Close the connection
mail.logout()

# Display additional details on the screen
print(f"Statistics for completed task for {USERNAME}:")
print(f"Date Range: from {start_date_str} to {end_date_str}")
print(f"Total Emails Processed: {total_emails}")
print(f"Total Unique Senders Fetched: {len(sender_data)}")

# Write the sender data to a CSV file
csv_file = f'fetched_senders_{USERNAME}_{start_date}-{end_date}.csv'

with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(['Sender Name', 'Sender Email Address'])
    for name, email_address in sender_data:
        writer.writerow([name, email_address])

print(f"Sender data has been saved to '{csv_file}'.")