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
        info["Name"] = name_element.find_next('td').get_text().strip()

    email_element = soup.find(string=re.compile(r'Email', re.IGNORECASE))
    if email_element:
        info["Email"] = email_element.find_next('td').get_text().strip()

    workshop_element = soup.find(string=re.compile(r'Workshop Detail', re.IGNORECASE))
    if workshop_element:
        info["Workshop Detail"] = workshop_element.find_next('td').get_text().strip()

    date_element = soup.find(string=re.compile(r'Date', re.IGNORECASE))
    if date_element:
        info["Date"] = date_element.find_next('td').get_text().strip()

    mobile_element = soup.find(string=re.compile(r'Mobile No\.', re.IGNORECASE))
    if mobile_element:
        info["Mobile No."] = mobile_element.find_next('td').get_text().strip()

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

if st.button("Fetch and Generate Excel"):
    try:
        # ... (existing code for email fetching)

        for num in mail_id_list:
            typ, data = my_mail.fetch(num, '(RFC822)')
            msg = email.message_from_bytes(data[0][1])

            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    html_content = part.get_payload(decode=True).decode('utf-8')
                    info = extract_info_from_html(html_content)

                elif part.get_content_type().startswith('image'):
                    image_bytes = part.get_payload(decode=True)
                    text = extract_text_from_image(image_bytes)
                    info["Image Text"] = text

                elif part.get_content_type() == 'application/msword':
                    docx_bytes = part.get_payload(decode=True)
                    text = extract_text_from_word(docx_bytes)
                    info["Word Text"] = text

                elif part.get_content_type() == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                    excel_bytes = part.get_payload(decode=True)
                    text = extract_text_from_excel(excel_bytes)
                    info["Excel Text"] = text

                elif part.get_content_type() == 'application/pdf':
                    pdf_bytes = part.get_payload(decode=True)
                    text = extract_text_from_pdf(pdf_bytes)
                    info["PDF Text"] = text

            # ... (existing code to add received date and append info to the list)

        # ... (existing code to create DataFrame, display, and download Excel file)

    except Exception as e:
        st.error(f"Error: {e}")
