import os
import json
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

CONFIG_FILE = 'config.json'
LOCAL_CONFIG = r'C:\Users\jlpms\OneDrive\Desktop\राजस्थान का इतिहास\Study_Automation\config.json'

def send_schedule_email(attachment_paths, recipient_email):
    # Try getting from Environment (GitHub)
    sender_email = os.environ.get("SENDER_EMAIL")
    sender_password = os.environ.get("SENDER_PASSWORD")
    recipient_email = os.environ.get("RECIPIENT_EMAIL", recipient_email)
    
    # Fallback to local config if not on Cloud
    config_to_use = CONFIG_FILE if os.path.exists(CONFIG_FILE) else (LOCAL_CONFIG if os.path.exists(LOCAL_CONFIG) else None)
    if not sender_email and config_to_use:
        with open(config_to_use, 'r') as f:
            config = json.load(f)
            sender_email = config.get("sender_email")
            sender_password = config.get("sender_password")
            recipient_email = config.get("recipient_email", recipient_email)

    if not sender_email or not sender_password:
        print("Error: Email credentials missing.")
        return

    # Create Message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"RAS Adaptive Study Plan - {datetime.now().strftime('%d %b')}"

    body = "Attached is your RAS Adaptive Study Plan and Today's PYQ Excel.\n\nKeep going! 'LEAD' (Learn, Engage, Adapt, Deliver)"
    msg.attach(MIMEText(body, 'plain'))

    # Attach Files
    for path in attachment_paths:
        if not path or not os.path.exists(path): continue
        with open(path, 'rb') as f:
            file_data = f.read()
            if path.endswith('.png'):
                part = MIMEImage(file_data, name=os.path.basename(path))
            else:
                from email.mime.application import MIMEApplication
                part = MIMEApplication(file_data, name=os.path.basename(path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(path)}"'
            msg.attach(part)

    # Send Email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        print(f"Email sent successfully to {recipient_email}")
    except Exception as e:
        print(f"Error sending email: {e}")
