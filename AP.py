import streamlit as st
import imaplib
import email
from datetime import datetime, timedelta

def download_attachments(username, password, target_email, start_date, attachment_format):
    # Convert start_date to datetime object
    start_date = datetime.strptime(start_date, '%Y-%m-%d')

    # Connect to the IMAP server (Gmail in this case)
    mail = imaplib.IMAP4_SSL('imap.gmail.com')

    # Login to your email account
    mail.login(username, password)

    # Select the mailbox (e.g., 'inbox')
    mail.select("inbox")

    # Calculate the date one day after the specified start_date
    one_day_after_start_date = start_date + timedelta(days=1)

    # Construct the search criterion using the date range and target email address
    search_criterion = f'(FROM "{target_email}" SINCE "{start_date.strftime("%d-%b-%Y")}" BEFORE "{one_day_after_start_date.strftime("%d-%b-%Y")}")'

    # Search for emails matching the criteria
    result, data = mail.search(None, search_criterion)

    # Iterate through the email IDs
    for num in data[0].split():
        result, msg_data = mail.fetch(num, "(RFC822)")
        raw_email = msg_data[0][1]

        # Parse the raw email content
        msg = email.message_from_bytes(raw_email)
        
        # Iterate through email parts
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            # Extract filename and content type
            filename = part.get_filename()
            content_type = part.get_content_type()

            # Check if the attachment format and file type match the user's choice
            if attachment_format.lower() in content_type.lower() and filename:
                # Decode and save the attachment
                file_data = part.get_payload(decode=True)
                with open(filename, 'wb') as f:
                    f.write(file_data)

    # Logout from the IMAP server
    mail.logout()

# Streamlit app
st.title("Email Attachment Downloader")

# Get user input through Streamlit
email_address = st.text_input("Enter your email address:")
password = st.text_input("Enter your email account password:", type="password")
target_email = st.text_input("Enter the email address from which you want to download attachments:")
start_date = st.text_input("Enter the start date (YYYY-MM-DD):")

# Dropdown for selecting attachment format
attachment_format = st.selectbox("Select the attachment format:", ["image", "pdf", "word", "excel"])

# Check if the user has provided all necessary inputs
if email_address and password and target_email and start_date and attachment_format:
    # Download attachments when the user clicks the button
    if st.button("Download Attachments"):
        # Download attachments
        download_attachments(email_address, password, target_email, start_date, attachment_format)
        st.success("Attachments downloaded successfully!")
else:
    st.warning("Please fill in all the required fields.")
