# Email Automation Project

## Overview
This Python script automates the process of sending personalized emails with attached letters to students using data from an Excel sheet. The emails are sent via a specified SMTP server, and the attachments are individualized for each student.

## Prerequisites
Before using the script, ensure you have the following:

- Python installed on your system.
- Required Python libraries: pandas, smtplib.
- An Excel file (`list.xlsx`) with a sheet named "Sheet1" containing student data, including columns for "Name" and "Email."

## Setup
1. Clone or download the project repository to your local machine.
2. Install the required Python libraries using the following command:
    ```
    pip install pandas smtplib
    ```
3. Place the Excel file (`list.xlsx`) with the student data in the project directory.
4. Create a folder named "attachments" in the project directory and place individualized PDF letters (named 1.pdf, 2.pdf, etc.) for each student in the folder.

## Usage
1. Run the script by executing the following command in the terminal:
    ```
    python mail.py
    ```
    or
    ```
    start.bat
    ```
    (if using Windows batch file)
2. Enter your email and password when prompted. This information is used to connect to the SMTP server.

## Configuration
- Update the `smtp_server`, `smtp_port`, `smtp_username`, and `smtp_password` variables with the appropriate values for your email provider.
- Customize the `sender_name` variable to reflect the sender's name in the emails.
- Adjust the `attachment_path` variable to point to the actual path of the PDF attachments.

## Important Notes
- Ensure that your email provider allows the use of less secure apps or generate an app password for authentication.
- Handle sensitive information, such as passwords, securely.

## Disclaimer
This script is provided as-is and may need modification based on specific use cases or email providers. Use it responsibly and comply with email service provider policies.

**Note:** It's crucial to be aware of the potential security risks associated with storing and transmitting sensitive information, such as email credentials. Consider more secure methods or token-based authentication for production use.