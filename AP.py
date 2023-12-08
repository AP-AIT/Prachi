import streamlit as st
import imaplib
import email
from datetime import datetime, timedelta
import io
from PyMuPDF import PyMuPDF

def display_pdfs(username, password, target_email, start_date):
    try:
        # Convert start_date to datetime object
        start_date = datetime.strptime(start_date, '%Y-%m-%d')

        # Connect to the IMAP server (Gmail in this case)
        mail = imaplib.IMAP4_SSL('imap.gmail.com')

        # Login to your email account
        mail.login(username, password)

        # Select the mailbox (e.g., 'inbox')
        mail.select("inbox")

        # Construct the search criterion using the date range and target email address
        search_criterion = f'(FROM "{target_email}" SINCE "{start_date.strftime("%d-%b-%Y")}" BEFORE "{(start_date + timedelta(days=1)).strftime("%d-%b-%Y")}")'

        # Search for emails matching the criteria
        result, data = mail.uid('search', None, search_criterion)
        email_ids = data[0].split()

        # List to store PDF data
        pdf_data = []

        # Iterate through the email IDs
        for email_id in email_ids:
            result, msg_data = mail.uid('fetch', email_id, "(RFC822)")
            raw_email = msg_data[0][1]

            # Parse the raw email content
            msg = email.message_from_bytes(raw_email)

            # Iterate through email parts
            for part in msg.walk():
                if part.get_content_type() == 'application/pdf':
                    # Extract PDF data
                    pdf_data.append(part.get_payload(decode=True))
    except Exception as e:
        st.error(f"An error occurred: {e}")
    finally:
        # Logout from the IMAP server (even if an error occurs)
        mail.logout()

    return pdf_data

# Streamlit app
st.title("PDF Viewer")

# Get user input through Streamlit
email_address = st.text_input("Enter your email address:")
password = st.text_input("Enter your email account password:", type="password")
target_email = st.text_input("Enter the email address from which you want to view PDFs:")
start_date = st.text_input("Enter the start date (YYYY-MM-DD):")

# Check if the user has provided all necessary inputs
if email_address and password and target_email and start_date:
    # Display PDFs when the user clicks the button
    if st.button("View PDFs"):
        # Display extracted PDFs
        pdf_data = display_pdfs(email_address, password, target_email, start_date)

        if not pdf_data:
            st.warning("No PDFs found.")

        for idx, pdf_bytes in enumerate(pdf_data, start=1):
            # Display PDF content using PyMuPDF
            pdf_doc = PyMuPDF.PdfDocument(io.BytesIO(pdf_bytes))
            pdf_text = ""
            for page_num in range(len(pdf_doc)):
                pdf_text += pdf_doc[page_num].getText()

            st.text_area(label=f'PDF {idx} Content', value=pdf_text, height=400)
else:
    st.warning("Please fill in all the required fields.")
