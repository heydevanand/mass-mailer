import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import getpass

# Function to send emails
def send_emails(smtp_server, smtp_port, smtp_username, smtp_password, df):
    from_email = smtp_username
    sender_name = 'Dy Registrar, Dr. Bhimrao Ambedkar University, Agra'

    iterator = 1

    for index, row in df.iterrows():
        to_email = row["Email"]
        recipient_name = row["Name"]

        subject = "RDC Letter"
        message = f"Dear {recipient_name},\n\nPlease find your RDC letter attached.\n\nRegards,\n{sender_name}"

        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(message.format(recipient_name), 'plain'))

        attachment_path = f'attachments/{iterator}.pdf'
        with open(attachment_path, 'rb') as attachment:
            part = MIMEApplication(attachment.read())
            part.add_header('Content-Disposition', 'attachment', filename=f'{recipient_name}.pdf')
            msg.attach(part)

        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(from_email, to_email, msg.as_string())
            iterator += 1
            print(f"Email sent to {recipient_name} ({to_email})")
        except Exception as e:
            print(f"Error sending email to {recipient_name}: {e}")

        if iterator > len(df):
            break

    server.quit()
    print("All emails sent successfully!")

if __name__ == '__main__':
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587  # Use 465 for SSL
    smtp_username = input("Enter your email: ")
    smtp_password = getpass.getpass("Enter your password: ")

    df = pd.read_excel('list.xlsx', sheet_name="Sheet1")

    # Call the function to send emails
    send_emails(smtp_server, smtp_port, smtp_username, smtp_password, df)
