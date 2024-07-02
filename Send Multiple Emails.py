import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def send_multiple_emails():
    # Set up email server
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    username = "joyaljoykarunagappally@gmail.com"
    password = "password"

    # Set up email content
    body_signature = "Sincerely,\nExcel Macro M"
    cc_address = "joyaljoy520@gmail.com"

    # Load data from Excel file (assuming it's in the same directory)
    import pandas as pd
    excel_file = "Details.xlsx"
    df = pd.read_excel(excel_file)

    # Set up email server connection
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(username, password)

    # Loop through each row in the Excel file
    for index, row in df.iterrows():
        # Create email message
        msg = MIMEMultipart()
        msg['To'] = row['C']
        msg['Subject'] = row['D']
        msg['CC'] = cc_address

        # Create email body
        body_header = f"Dear {row['B']},"
        body_main = row['E']
        body = f"{body_header}\n\n{body_main}\n\n{body_signature}"

        # Add attachment (assuming it's a file path in column F)

        attachment = row['F']
        filename = attachment
        with open(attachment, 'rb') as f:
            file_content = f.read()
        msg.attach(MIMEText(file_content, 'octet-stream'))
        msg['Content-Disposition'] = f'attachment; filename="{filename}"'

        # Send email
        server.sendmail(username, row['C'], msg.as_string())

    # Close email server connection
    server.quit()


send_multiple_emails()
