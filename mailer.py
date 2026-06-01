import os
import json
import smtplib
from datetime import datetime, timedelta
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
    
    tomorrow = datetime.now() + timedelta(days=1)
    is_weekend_plan = (tomorrow.weekday() == 5)
    
    if custom_subject and ("Weekend" in custom_subject or "वीकेंड" in custom_subject or "Weekend Planner" in custom_subject):
        is_weekend_plan = True

    if is_weekend_plan:
        sat_date = tomorrow
        sun_date = tomorrow + timedelta(days=1)
        if sat_date.month == sun_date.month:
            weekend_range = f"{sat_date.strftime('%d')} & {sun_date.strftime('%d %B %Y')}"
        else:
            weekend_range = f"{sat_date.strftime('%d %B')} & {sun_date.strftime('%d %B %Y')}"

        if custom_subject:
            msg['Subject'] = custom_subject
        else:
            msg['Subject'] = f"📅 Weekend Planner - {weekend_range} (शनिवार एवं रविवार)"
            
        html_body = f"""
        <html>
        <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.6; color: #1e293b; background-color: #f8fafc; padding: 20px; margin: 0;">
            <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 30px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
                <div style="font-size: 15px; font-weight: 600; color: #475569; margin-bottom: 12px;">
                    RAS Aspirant,
                </div>
                <div style="font-size: 18px; font-weight: 700; color: #1e3a8a; margin-bottom: 16px;">
                    📅 आपका Weekend Planner ({weekend_range}) का प्लानर तैयार है
                </div>
                <p style="font-size: 15px; color: #334155; margin: 0 0 16px 0;">
                    आपका Weekend Planner ({weekend_range}) डिज़ाइन प्रीव्यू तैयार है और संलग्न है।
                </p>
                <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 16px; margin: 20px 0;">
                    <ul style="margin: 0; padding-left: 0; font-size: 15px; color: #166534; list-style-type: none;">
                        <li style="margin-bottom: 8px;">✅ विषयवार अध्ययन योजना</li>
                        <li style="margin-bottom: 8px;">✅ रिवीजन एवं प्रैक्टिस स्लॉट</li>
                        <li style="margin-bottom: 8px;">✅ लक्ष्य आधारित अध्ययन कार्यक्रम</li>
                        <li style="margin-bottom: 0;">✅ बेहतर समय प्रबंधन</li>
                    </ul>
                </div>
                <p style="font-size: 15px; color: #334155; margin: 0 0 20px 0;">
                    कृपया प्लान का अवलोकन करें और आगामी वीकेंड की तैयारी आज ही सुनिश्चित करें।
                </p>
                <div style="margin: 24px 0; padding: 16px; background-color: #eff6ff; border-left: 4px solid #3b82f6; border-radius: 6px; font-size: 15px; font-weight: 600; color: #1e40af;">
                    हर योजनाबद्ध वीकेंड आपको अपने अंतिम लक्ष्य के एक कदम और करीब ले जाता है। {extra_msg}
                </div>
                <p style="font-size: 14px; font-weight: 600; color: #64748b; margin: 24px 0 0 0; border-top: 1px solid #e2e8f0; padding-top: 16px;">
                    Automation System 🚀
                </p>
            </div>
        </body>
        </html>
        """
    else:
        if custom_subject:
            msg['Subject'] = custom_subject
        else:
            msg['Subject'] = f"RAS Adaptive Study Plan - {datetime.now().strftime('%d %b %Y')}"

        html_body = f"""
        <html>
        <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.6; color: #1e293b; background-color: #f8fafc; padding: 20px; margin: 0;">
            <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 30px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
                <div style="font-size: 18px; font-weight: 700; color: #1e3a8a; margin-bottom: 16px;">
                    📅 कल का RAS Study Plan उपलब्ध है
                </div>
                <p style="font-size: 15px; color: #334155; margin: 0 0 16px 0;">
                    आपका कल का RAS Study Plan तैयार कर दिया गया है और संलग्न है।
                </p>
                <div style="margin: 24px 0; padding: 16px; background-color: #fef2f2; border-left: 4px solid #ef4444; border-radius: 6px; font-size: 15px; font-weight: 600; color: #991b1b;">
                    योजनाबद्ध तैयारी + निरंतर प्रयास = RAS में सफलता {extra_msg}
                </div>
                <p style="font-size: 14px; font-weight: 600; color: #64748b; margin: 24px 0 0 0; border-top: 1px solid #e2e8f0; padding-top: 16px;">
                    Automation System 🚀
                </p>
            </div>
        </body>
        </html>
        """
    msg.attach(MIMEText(html_body, 'html'))

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
