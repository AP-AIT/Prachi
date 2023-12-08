import streamlit as st
import imaplib
import email
from datetime import datetime, timedelta
import io
import fitz  # PyMuPDF

def extract_pdfs(username, password, target_email, start_date):
    try:
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(username, password)
        mail.select("inbox")
        search_criterion = f'(FROM "{target_email}" SINCE "{start_date.strftime("%d-%b-%Y")}" BEFORE "{(start_date + timedelta(days=1)).strftime("%d-%b-%Y")}")'
        result, data = mail.uid('search', None, search_criterion)
        email_ids = data[0].split()
        pdf_data = []
        for email_id in email_ids:
            result, msg_data = mail.uid('fetch', email_id, "(RFC822)")
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            for part in msg.walk():
                if part.get_content_maintype() == 'application' and part.get_content_subtype() == 'pdf':
                    pdf_data.append(part.get_payload(decode=True))
    except imaplib.IMAP4_SSL.error as e:
        st.error(f"IMAP error occurred: {e}")
    except email.errors.MessageError as e:
        st.error(f"Message error occurred: {e}")
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
    finally:
        mail.logout()
    return pdf_data

# Streamlit app
st.title("PDF Extractor")

email_address = st.text_input("Enter your email address:")
password = st.text_input("Enter your email account password:", type="password")
target_email = st.text_input("Enter the email address from which you want to extract PDFs:")
start_date = st.text_input("Enter the start date (YYYY-MM-DD):")

if email_address and password and target_email and start_date:
    if st.button("Extract PDFs"):
        st.info("Extracting PDFs. Please wait...")
        pdf_data = extract_pdfs(email_address, password, target_email, start_date)
        if not pdf_data:
            st.warning("No PDFs found.")
        for idx, pdf_bytes in enumerate(pdf_data, start=1):
            pdf_document = fitz.open(stream=io.BytesIO(pdf_bytes), filetype="pdf")
            st.write(f'PDF {idx} - Number of pages: {pdf_document.page_count}')
        st.success("Extraction complete!")
else:
    st.warning("Please fill in all the required fields.")
