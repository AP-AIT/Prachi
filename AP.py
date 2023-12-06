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
import pytesseract

# ... (other imports and functions)

# Streamlit app title
st.title("Automate2Excel: Simplified Data Transfer")

# ... (other input fields and functions)

# Initialize mail_id_list outside the try block
mail_id_list = []

if st.button("Fetch and Generate Extracted Document"):
    try:
        # ... (existing code for email fetching)

        info_list = []

        # Iterate through messages and extract information based on the selected document type
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

                elif document_type == "Image" and part.get_content_type().startswith("image"):
                    # Call the function to extract text from image
                    image_text = extract_text_from_image(part.get_payload(decode=True))
                    st.text(f"Extracted Text from Image: {image_text}")

                elif document_type == "Word" and part.get_content_type().startswith("application/msword"):
                    # Call the function to extract text from Word document
                    word_text = extract_text_from_word(part.get_payload(decode=True))
                    st.text(f"Extracted Text from Word: {word_text}")

                elif document_type == "Excel" and part.get_content_type().startswith("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
                    # Call the function to extract text from Excel file
                    excel_text = extract_text_from_excel(part.get_payload(decode=True))
                    st.text(f"Extracted Text from Excel: {excel_text}")

                elif document_type == "PDF" and part.get_content_type().startswith("application/pdf"):
                    # Call the function to extract text from PDF
                    pdf_text = extract_text_from_pdf(part.get_payload(decode=True))
                    st.text(f"Extracted Text from PDF: {pdf_text}")

        # ... (existing code for DataFrame creation and display)

        st.success("Extraction completed.")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.warning("Click the 'Fetch and Generate Extracted Document' button to retrieve and process emails.")
