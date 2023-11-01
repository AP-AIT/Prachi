import streamlit as st
st.header("MKSSS-AIT")
import imaplib
import email
from bs4 import BeautifulSoup
import re
import pandas as pd

# Streamlit app title
st.title("Email Data to Excel App")

# Create input fields for the user and password
user = st.text_input("Enter your email address")
password = st.text_input("Enter your email password", type="password")

# Function to extract information from HTML content
def extract_info_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    name_element = soup.find(string=re.compile(r'Name', re.IGNORECASE))
    email_element = soup.find(string=re.compile(r'Email', re.IGNORECASE))
    workshop_element = soup.find(string=re.compile(r'Workshop Detail', re.IGNORECASE))
    date_element = soup.find(string=re.compile(r'Date', re.IGNORECASE))
    mobile_element = soup.find(string=re.compile(r'Mobile No\.', re.IGNORECASE))

    info = {
        "Name": None,
        "Email": None,
        "Workshop Detail": None,
        "Date": None,
        "Mobile No.": None
    }

    if name_element:
        info["Name"] = name_element.find_next('td').get_text().strip()

    if email_element:
        info["Email"] = email_element.find_next('td').get_text().strip()

    if workshop_element:
        info["Workshop Detail"] = workshop_element.find_next('td').get_text().strip()

    if date_element:
        info["Date"] = date_element.find_next('td').get_text().strip()

    if mobile_element:
        info["Mobile No."] = mobile_element.find_next('td').get_text().strip()

    return info

if st.button("Fetch and Generate Excel"):
    try:
        # URL for IMAP connection
        imap_url = 'imap.gmail.com'

        # Connection with GMAIL using SSL
        my_mail = imaplib.IMAP4_SSL(imap_url)

        # Log in using user and password
        my_mail.login(user, password)

        # Select the Inbox to fetch messages
        my_mail.select('inbox')

        # Define the key and value for email search
        key = 'FROM'
        value = 'info@mkssscareerguidanceexpo.com'
        _, data = my_mail.search(None, key, value)

        mail_id_list = data[0].split()

        info_list = []

        # Iterate through messages and extract information from HTML content
        for num in mail_id_list:
            typ, data = my_mail.fetch(num, '(RFC822)')
            msg = email.message_from_bytes(data[0][1])

            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    html_content = part.get_payload(decode=True).decode('utf-8')
                    info = extract_info_from_html(html_content)

                    # Extract and add the received date
                    date = msg["Date"]
                    info["Received Date"] = date

                    info_list.append(info)

        # Create a DataFrame from the info_list
        df = pd.DataFrame(info_list)

        # Generate the Excel file
        st.write("Data extracted from emails:")
        st.write(df)
        st.write("Downloading Excel file...")
        df.to_excel('EXPO_leads.xlsx', index=False, engine='openpyxl')
        st.success("Excel file has been generated and is ready for download.")

    except Exception as e:
        st.error(f"Error: {e}")

# ... (The rest of your code)