import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Read student data from the Excel sheet
df = pd.read_excel('list.xlsx', sheet_name="Sheet1")

# Set up your email server and credentials
smtp_server = 'smtp.gmail.com'
smtp_port = 587  # Use 465 for SSL
smtp_username = input("Enter your email: ")
smtp_password = input("Enter your password: ")
from_email = smtp_username
sender_name = 'Dy Registrar, Dr. Bhimrao Ambedkar University, Agra'

# Iterate through the student data and send emails
for index, row in df.iterrows():
    to_email = row["Email"]
    recipient_name = row["Name"]

    # Email subject and message
    subject = "RDC Letter"
    message = f"Dear {recipient_name},\n\nPlease find your RDC letter attached.\n\nRegards,\n{sender_name}"

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message.format(recipient_name), 'plain'))

    iterator = 1

    # Iterate through the student data and send emails
    for index, row in df.iterrows():
        to_email = row["Email"]
        recipient_name = row["Name"]

        # Email subject and message
        subject = "RDC Letter"
        message = f"Dear {recipient_name},\n\nPlease find your RDC letter attached.\n\nRegards,\n{sender_name}"

        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(message.format(recipient_name), 'plain'))

        # Attach the file
        attachment_path = f'attachments/{iterator}.pdf'  # Replace with the actual path to the attachment
        with open(attachment_path, 'rb') as attachment:
            part = MIMEApplication(attachment.read())
            part.add_header('Content-Disposition', 'attachment', filename=f'{recipient_name}.pdf')
            msg.attach(part)

        # Connect to the SMTP server and send the email
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(from_email, to_email, msg.as_string())
            iterator += 1
            print(f"Email sent to {recipient_name} ({to_email})")
        except Exception as e:
            print(f"Error sending email to {recipient_name}: {e}")

        # Break out of the loop if there are no more files
        if iterator > len(df):
            break

    # Break out of the loop if there are no more files
    if iterator > len(df):
        break

        

# Disconnect from the server
server.quit()

print("All emails sent successfully!")