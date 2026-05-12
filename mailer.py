import os
import json
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

CONFIG_FILE = 'config.json'
POSSIBLE_CONFIGS = [
    'config.json',
    r'R:\Study_Automation_System\config.json',
    r'C:\Users\jlpms\OneDrive\Desktop\राजस्थान का इतिहास\Study_Automation\config.json'
]

def send_schedule_email(attachment_paths, recipient_email, extra_msg=""):
    # ... (rest of function signature and setup)
    sender_email = os.environ.get("SENDER_EMAIL")
    sender_password = os.environ.get("SENDER_PASSWORD")
    recipient_email = os.environ.get("RECIPIENT_EMAIL", recipient_email)
    
    if not recipient_email or "@" not in str(recipient_email):
        recipient_email = "figuring.cse@gmail.com"
    
    # Fallback to local config if not on Cloud
    config_to_use = None
    for p in POSSIBLE_CONFIGS:
        if os.path.exists(p):
            config_to_use = p
            break
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

    body = f"Attached is your RAS Adaptive Study Plan and Today's PYQ Excel.\n\nKeep going! 'LEAD' (Learn, Engage, Adapt, Deliver){extra_msg}"
    msg.attach(MIMEText(body, 'plain'))

    # Attach Files
    print(f"Total attachments to process: {len(attachment_paths)}")
    for path in attachment_paths:
        if not path:
            print("Skipping empty attachment path")
            continue
        if not os.path.exists(path):
            print(f"Error: Attachment file not found at {os.path.abspath(path)}")
            continue
        
        print(f"Attaching file: {path}")
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
        print(f"Connecting to SMTP server...")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        print(f"✅ Email successfully sent to {recipient_email}")
    except Exception as e:
        print(f"❌ SMTP Error: {e}")
        if "Authentication failed" in str(e):
            print("   TIP: Check your App Password in config.json. It must be a 16-character code.")
        raise e
