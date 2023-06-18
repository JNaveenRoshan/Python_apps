import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Read the Excel file
df = pd.read_excel('data.xlsx')

# Email credentials
sender_email = '<your mail address>'
sender_password = '<App password>'

# Email template
subject_template = 'Greetings, {first_name}!'
body_template = '''
Hi, {first_name},



Check the email please


'''

# SMTP server setup
smtp_server = 'smtp.gmail.com'
smtp_port = 587

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    # Retrieve name and email address
    name = row['Name']
    email = row['Email Address']
    first_name = name.split()[0]  # Extract the first name from the full name

    # Construct the subject and email body by replacing the template placeholders
    subject = subject_template.format(first_name=first_name)
    body = body_template.format(first_name=first_name)

    # Create the email message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = email
    message['Cc'] = '<the mail that needs to be in CC>' 
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    # Connect to the SMTP server and send the email
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(message)

    print(f"Email sent to {name} at {email}")

print("All emails sent successfully.")
