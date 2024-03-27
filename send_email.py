import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
def send_email_via_outlook(sender_email, password, recipient_email, subject, body, attachment_paths):
    # SMTP server configuration for Outlook.com
    smtp_server = "smtp.office365.com"
    smtp_port = 587  # TLS

    # Create MIME message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))

    # Process any attachments
    if len(attachment_paths)>0:
        for path in attachment_paths:
            part = MIMEBase('application', 'octet-stream')
            with open(path, 'rb') as file:
                part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={path.split('/')[2]}')
            msg.attach(part)

    try:
        # Connect to the SMTP server using TLS
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        
        # Login to the SMTP server
        server.login(sender_email, password)
        
        # Send the email
        server.send_message(msg)
        
        # Close the SMTP server connection
        server.quit()
        
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")
