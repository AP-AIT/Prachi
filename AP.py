import streamlit as st
import imaplib
import email
from datetime import datetime, timedelta
import io
from PIL import Image
import pytesseract

# Specify the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = "C:\Program Files\Tesseract-OCR" # Replace with the actual path

def extract_text_from_images(image_data):
    extracted_text = []
    for image_bytes in image_data:
        # Use pytesseract to extract text from image
        image = Image.open(io.BytesIO(image_bytes))
        text = pytesseract.image_to_string(image)
        extracted_text.append(text.strip())
    return extracted_text

def display_images(username, password, target_email, start_date):
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

    # List to store image data
    image_data = []

    # Iterate through the email IDs
    for num in data[0].split():
        result, msg_data = mail.fetch(num, "(RFC822)")
        raw_email = msg_data[0][1]

        # Parse the raw email content
        msg = email.message_from_bytes(raw_email)
        
        # Iterate through email parts
        for part in msg.walk():
            if part.get_content_maintype() == 'image':
                # Extract image data
                image_data.append(part.get_payload(decode=True))

    # Logout from the IMAP server
    mail.logout()

    return image_data

# Streamlit app
st.title("Image Text Extractor")

# Get user input through Streamlit
email_address = st.text_input("Enter your email address:")
password = st.text_input("Enter your email account password:", type="password")
target_email = st.text_input("Enter the email address from which you want to extract text from images:")
start_date = st.text_input("Enter the start date (YYYY-MM-DD):")

# Check if the user has provided all necessary inputs
if email_address and password and target_email and start_date:
    # Extract text from images when the user clicks the button
    if st.button("Extract Text from Images"):
        # Display extracted images
        image_data = display_images(email_address, password, target_email, start_date)
        
        # Extract text from images
        extracted_text = extract_text_from_images(image_data)
        
        # Display extracted text
        st.write("Extracted Text from Images:")
        for text in extracted_text:
            st.write(text)

        # Download text when the user clicks the button
        if st.button("Download Text"):
            text_filename = "extracted_text.txt"
            with open(text_filename, 'w') as file:
                for text in extracted_text:
                    file.write(text + "\n")
            st.success(f"Text downloaded successfully. Filename: {text_filename}")
else:
    st.warning("Please fill in all the required fields.")
