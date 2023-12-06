import streamlit as st
import imaplib
import email
from bs4 import BeautifulSoup
import re
import pandas as pd
import base64
from PIL import Image
import io
import docx2txt
import openpyxl
import PyPDF2

# Streamlit app title
st.title("Automate2Excel: Simplified Data Transfer")

# Create input fields for the user and password
user = st.text_input("Enter your email address")
password = st.text_input("Enter your email password", type="password")

# Create input field for the email address to search for
search_email = st.text_input("Enter the email address to search for")

# Function to extract information from HTML content
def extract_info_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    info = {
        "Name": None,
        "Email": None,
        "Workshop Detail": None,
        "Date": None,
        "Mobile No.": None
    }

    name_element = soup.find(string=re.compile(r'Name', re.IGNORECASE))
    if name_element:
        next_td = name_element.find_next('td')
        info["Name"] = next_td.get_text().strip() if next_td else None

    email_element = soup.find(string=re.compile(r'Email', re.IGNORECASE))
    if email_element:
        next_td = email_element.find_next('td')
        info["Email"] = next_td.get_text().strip() if next_td else None

    workshop_element = soup.find(string=re.compile(r'Workshop Detail', re.IGNORECASE))
    if workshop_element:
        next_td = workshop_element.find_next('td')
        info["Workshop Detail"] = next_td.get_text().strip() if next_td else None

    date_element = soup.find(string=re.compile(r'Date', re.IGNORECASE))
    if date_element:
        next_td = date_element.find_next('td')
        info["Date"] = next_td.get_text().strip() if next_td else None

    mobile_element = soup.find(string=re.compile(r'Mobile No\.', re.IGNORECASE))
    if mobile_element:
        next_td = mobile_element.find_next('td')
        info["Mobile No."] = next_td.get_text().strip() if next_td else None

    return info

# Function to extract text from an image
def extract_text_from_image(image_bytes):
    image = Image.open(io.BytesIO(image_bytes))
    text = pytesseract.image_to_string(image)
    return text

# Function to extract text from Word document
def extract_text_from_word(docx_bytes):
    text = docx2txt.process(io.BytesIO(docx_bytes))
    return text

# Function to extract text from Excel file
def extract_text_from_excel(excel_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    text = []
    for sheet in wb.sheetnames:
        for row in wb[sheet].iter_rows(values_only=True):
            text.extend(row)
    return text

# Function to extract text from PDF
def extract_text_from_pdf(pdf_bytes):
    pdf_reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    text = ""
    for page_num in range(len(pdf_reader.pages)):
        text += pdf_reader.pages[page_num].extract_text()
    return text

# Initialize mail_id_list outside the try block
mail_id_list = []

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
        value = search_email  # Use the user-inputted email address to search
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

                # ... (existing code for other attachment types)

        # Create a DataFrame from the info_list
        df = pd.DataFrame(info_list)

        # Generate the Excel file
        st.write("Data extracted from emails:")
        st.write(df)

        if st.button("Download Excel File"):
            excel_file = df.to_excel('EXPO_leads.xlsx', index=False, engine='openpyxl')
            if excel_file:
                with open('EXPO_leads.xlsx', 'rb') as file:
                    st.download_button(
                        label="Click to download Excel file",
                        data=file,
                        key='download-excel'
                    )

        st.success("Excel file has been generated and is ready for download.")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.warning("Click the 'Fetch and Generate Excel' button to retrieve and process emails.")
