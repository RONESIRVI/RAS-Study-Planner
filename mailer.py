import os
import json
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

CONFIG_FILE = 'config.json'
POSSIBLE_CONFIGS = [
    'config.json',
    os.path.join('data', 'config.json'),
    r'R:\Study_Automation_System\config.json'
]

def send_schedule_email(attachment_paths, recipient_email, extra_msg="", custom_subject=None):
    # ... (Environment variables logic) ...
    sender_email = os.environ.get("SENDER_EMAIL")
    sender_password = os.environ.get("SENDER_PASSWORD")
    recipient_email = os.environ.get("RECIPIENT_EMAIL", recipient_email)

    # Fallback to local config
    if not sender_email or not sender_password:
        config_to_use = None
        for p in POSSIBLE_CONFIGS:
            if os.path.exists(p):
                config_to_use = p
                break
        if config_to_use:
            with open(config_to_use, 'r') as f:
                config = json.load(f)
                sender_email = sender_email or config.get("sender_email")
                sender_password = sender_password or config.get("sender_password")
                recipient_email = recipient_email or config.get("recipient_email")

    if not sender_email or not sender_password:
        print("[ERROR] Email credentials missing. (Check GitHub Secrets or config.json)")
        return

    # Create Message
    msg = MIMEMultipart()
    msg['From'] = f"RAS Mentorship System <{sender_email}>"
    msg['To'] = recipient_email
    
    if custom_subject:
        msg['Subject'] = custom_subject
    else:
        msg['Subject'] = f"RAS Adaptive Study Plan - {datetime.now().strftime('%d %b %Y')}"

    body = f"📅 कल का RAS Study Plan उपलब्ध है\n\nआपका कल का RAS Study Plan तैयार कर दिया गया है और संलग्न है।\n\nयोजनाबद्ध तैयारी + निरंतर प्रयास = RAS में सफलता {extra_msg}\n\nAutomation System 🚀"
    msg.attach(MIMEText(body, 'plain'))

    # Attach Files
    for path in attachment_paths:
        if path and os.path.exists(path):
            print(f"Attaching: {path}")
            with open(path, 'rb') as f:
                file_data = f.read()
                name = os.path.basename(path)
                if path.endswith('.png'):
                    part = MIMEImage(file_data, name=name)
                else:
                    part = MIMEApplication(file_data, Name=name)
                    part['Content-Disposition'] = f'attachment; filename="{name}"'
                msg.attach(part)

    # Send Email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        print(f"[SUCCESS] Email successfully sent to {recipient_email}")
    except Exception as e:
        print(f"[ERROR] SMTP Error: {e}")

def send_email(attachments, extra_msg="", custom_subject=None):
    # Main entry point for automation_main.py
    send_schedule_email(attachments, "figuring.cse@gmail.com", extra_msg, custom_subject)
