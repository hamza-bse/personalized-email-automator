import os
from dotenv import load_dotenv
import pandas as pd
import smtplib
from email.message import EmailMessage

# Load credentials from .env file
load_dotenv()
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Load contacts from Excel file
file_path = "C:/Users/Hamza Kamal/Desktop/python_project.xlsx"
contacts = pd.read_excel(file_path)

# Set up the SMTP server
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

# Send personalized emails
for index, row in contacts.iterrows():
    name = row["Name"]
    recipient_email = row["email"]

    if pd.isna(recipient_email) or pd.isna(name):
        print(f"Skipping row {index+1} due to missing data")
        continue

    msg = EmailMessage()
    msg["Subject"] = "A Personal Note for You"
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = recipient_email
    msg.set_content(f"""\

Hi {name},

Hope you're doing well!

This is a personalized email sent from my Python program. If you have any questions, feel free to reach out.

Best regards,  
Hamza Kamal
""")

    try:
        server.send_message(msg)
        print(f"Email sent to {name} at {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {name}: {e}")

server.quit()
