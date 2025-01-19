import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime

# Email configuration
smtp_server = "smtp.gmail.com"  # Gmail SMTP server
smtp_port = 587  # Port for TLS
sender_email = "jaydip.comline@gmail.com"
sender_password = "ibzn qbwz ejvl voag"  # Replace with your app password
recipients = ["parthiv.kansara@comline.uk.com", "hardik.goswami@comline.uk.com", "jaydip.d@comline.uk.com"]

# File directory configuration
file_path = r"C:\Users\jaydi\OneDrive - Comline\Daily_Stock_Report\FINAL REPORT TO SEND TO ROBERT"
file_prefix = "REPLEN_FORMAT_"  # Prefix of the file name

# Get the file name dynamically based on date and time pattern
def get_latest_file(file_path, file_prefix):
    try:
        # List all files in the directory with the specified prefix
        files = [f for f in os.listdir(file_path) if f.startswith(file_prefix) and f.endswith(".xlsx")]
        if not files:
            return None  # No matching files found
        
        # Sort files by modification time
        files = sorted(files, key=lambda f: os.path.getmtime(os.path.join(file_path, f)), reverse=True)
        return files[0]  # Return the most recently modified file
    except Exception as e:
        print(f"Error finding the latest file: {e}")
        return None

# Email subject and body
subject = "Daily Replen Format Report"
body = """
Dear Team,

Please find attached the Replen Format report for today.

Best regards,
Jaydip
"""

# Function to send email with an attachment
def send_email_with_attachment(sender, recipients, subject, body, file_path, smtp_server, smtp_port, sender_password):
    try:
        # Create the email object
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = ", ".join(recipients)
        msg['Subject'] = subject

        # Attach the email body
        msg.attach(MIMEText(body, 'plain'))

        # Attach the file
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(file_path)}",
        )
        msg.attach(part)

        # Send the email
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Start TLS encryption
        server.login(sender, sender_password)
        server.send_message(msg)
        server.quit()

        print(f"Email sent successfully to: {', '.join(recipients)}")
    except Exception as e:
        print(f"Error while sending email: {e}")

# Find the latest file and send the email
latest_file = get_latest_file(file_path, file_prefix)
if latest_file:
    full_file_path = os.path.join(file_path, latest_file)
    print(f"Sending email with file: {full_file_path}")
    send_email_with_attachment(
        sender=sender_email,
        recipients=recipients,
        subject=subject,
        body=body,
        file_path=full_file_path,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        sender_password=sender_password
    )
else:
    print(f"No file found with prefix '{file_prefix}' in directory: {file_path}")
