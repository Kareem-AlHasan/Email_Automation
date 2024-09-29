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
    msg['Subject'] = "Interreg NEXT MED: Proposal for Partnership in Jordan"
    
    # Email body
    body = f"Dear {recipient_name},\n\nI hope this email finds you well. My name is Kareem, and I am reaching out to you on behalf of LIVINC Jordan, a Creative Community Incubator nestled in the Heart of Nature. We are dedicated to fostering a thriving ecosystem that empowers talents in the Creative Economy to achieve sustainable development.\n\nGiven our shared interests in smart solutions, green economy, and social projects, we believe that collaborating on initiatives such as Interreg NEXT MED could yield significant benefits for both parties.\n\nAt LIVINC Jordan, we offer a range of programs and services tailored to support and nurture creative individuals and projects:\n\n1. ACADEMY: What they don't teach at school!\n2. CREATIVE INCUBATOR: The environment to create and grow!\n3. MARKETPLACE: The showcase channels for your work!\n\nWe understand that the deadline for partnership proposals for Interreg NEXT MED is fast approaching on May 30th. Despite the short notice, we are fully committed to moving swiftly to explore the opportunity to be your Jordan partner.\n\nAttached to this email, you will find a detailed profile of LIVINC Jordan, which provides further insights into our mission, values, and offerings. Additionally, we are more than happy to provide any additional information or materials that you may require to consider our proposal.\n\nWe are genuinely excited about the possibility of partnering with your esteemed organization and are eager to discuss this opportunity further. Please let us know a convenient time for a meeting or call to explore how we can collaborate effectively.\n\nThank you for considering our proposal. We look forward to the possibility of working together to drive positive change and innovation in our shared areas of interest.\n\nWarm regards,\n\nKareem Shadi\nPartnership Manager\nLIVINC Jordan"
    msg.attach(MIMEText(body, 'plain'))
    
    # Attachment
    filename = "LIVINC Incubator.pdf"
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
