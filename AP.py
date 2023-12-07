import streamlit as st
import imaplib
import email
from datetime import datetime, timedelta
import io
from PIL import Image

def display_images(username, password, target_email, start_date):
    try:
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
        search_criterion = f'(FROM "{target_email}" SENTSINCE "{start_date} 00:00:00" BEFORE "{one_day_after_start_date} 00:00:00")'

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
    except Exception as e:
        st.error(f"An error occurred: {e}")
    finally:
        # Logout from the IMAP server (even if an error occurs)
        mail.logout()

    return image_data

# Streamlit app
st.title("Image Viewer")

# Get user input through Streamlit
email_address = st.text_input("Enter your email address:")
password = st.text_input("Enter your email account password:", type="password")
target_email = st.text_input("Enter the email address from which you want to view images:")
start_date = st.text_input("Enter the start date (YYYY-MM-DD):")

# Check if the user has provided all necessary inputs
if email_address and password and target_email and start_date:
    # Display images when the user clicks the button
    if st.button("View Images"):
        # Display extracted images
        image_data = display_images(email_address, password, target_email, start_date)

        if not image_data:
            st.warning("No images found.")

        for idx, image_bytes in enumerate(image_data, start=1):
            # Display image using PIL
            image = Image.open(io.BytesIO(image_bytes))
            st.image(image, caption=f'Image {idx}', use_column_width=True)
else:
    st.warning("Please fill in all the required fields.")
