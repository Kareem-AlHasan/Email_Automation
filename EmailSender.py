import random
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time

# Read the Excel file
df = pd.read_excel('recipients.xlsx')

# Email credentials
sender_email = ""
password = "" 

# Email server configuration
smtp_server = "smtp.gmail.com"
smtp_port = 587

# Create the email server connection
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(sender_email, password)
# Loop through the Excel file and send emails
for index, row in df.iterrows():
    recipient_name = row['Name']
    recipient_email = row['Email']

    
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Subject..."
    
    # Email body
    body = f"Dear {recipient_name},\n\nI hope this email finds you well. My name is..."
    msg.attach(MIMEText(body, 'plain'))
    
    # Attachment
    filename = "File.pdf"
    attachment = open(filename, "rb")
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {filename}")
    msg.attach(part)
    
    # Send the email
    server.sendmail(sender_email, [recipient_email], msg.as_string())
    print(f"Email sent to {recipient_name} at {recipient_email}")
    # random number of seconds to sleep
    time.sleep(random.randint(1, 4))

# Quit the server connection
server.quit()
