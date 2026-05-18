import smtplib
import json
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# Load configuration
def load_config():
    config_path = "config.json"
    if not os.path.exists(config_path):
        config_path = os.path.join("R:\\Study_Automation_System", "config.json")
    
    with open(config_path, 'r') as f:
        return json.load(f)

def send_evening_reminder():
    config = load_config()
    now = datetime.now()
    current_day = now.strftime('%A')
    is_weekend = current_day in ['Saturday', 'Sunday']
    
    if is_weekend:
        scheduled_time = config.get("weekend_evening_time", "13:00")
    else:
        scheduled_time = config.get("weekday_evening_time", "13:00")
        
    weekend_mode = config.get("weekend_mode", True)

    print(f"--- Starting Evening Reminder at {now.strftime('%Y-%m-%d %H:%M:%S')} ---")
    print(f"Scheduled Time: {scheduled_time} | Weekend Mode: {'ON' if weekend_mode else 'OFF'}")

    msg = MIMEMultipart()
    msg['From'] = config['sender_email']
    msg['To'] = config['recipient_email']
    
    if weekend_mode and is_weekend:
        msg['Subject'] = f"🕯️ Weekend Reflection - {now.strftime('%d %b %Y')}"
        title = "Weekend Reflection"
        message_body = """
        <p>नमस्ते,</p>
        <p>आज वीकेंड है, आशा है कि आपने अपनी पढ़ाई का रिवीज़न (Revision) अच्छे से किया होगा।</p>
        <p><b>वीकेंड टास्क:</b></p>
        <p>1. पिछले सप्ताह के छूटे हुए टॉपिक्स (Backlogs) को चेक करें।<br>
        2. अगले सप्ताह के लिए अपनी ऊर्जा संचित करें।<br>
        3. अपनी <b>Master Tracker Sheet</b> में आज के कार्यों का स्टेटस (उन्हें <b>'done'</b> मार्क करना न भूलें) अपडेट कर दें।</p>
        <div style="margin: 30px 0; padding: 20px; background: #f8fafc; border-left: 4px solid #f59e0b; border-radius: 4px;">
            <i>"तैयारी करने में विफल रहने का मतलब है विफल होने की तैयारी करना।"</i>
        </div>
        """
    else:
        msg['Subject'] = f"📊 Daily Progress Report - {now.strftime('%d %b %Y')}"
        title = "Daily Milestone Review"
        message_body = """
        <p>नमस्ते,</p>
        <p>आशा है कि आज का आपका अध्ययन (Study Session) फलदायी रहा होगा।</p>
        <p>कृपया अपनी <b>Master Tracker Sheet</b> में आज के कार्यों का स्टेटस (उन्हें <b>'done'</b> मार्क करना न भूलें.) अपडेट कर दें।</p>
        <div style="margin: 30px 0; padding: 20px; background: #f8fafc; border-left: 4px solid #00d2ff; border-radius: 4px;">
            <i>"अनुशासन ही लक्ष्य और उपलब्धि के बीच का पुल है।"</i>
        </div>
        """

    # Professional HTML Body
    html_content = f"""
    <html>
    <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 10px; overflow: hidden;">
            <div style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 30px; text-align: center; color: white;">
                <h1 style="margin: 0; font-size: 24px; letter-spacing: 1px;">{title}</h1>
            </div>
            <div style="padding: 40px; background: #ffffff;">
                {message_body}
                <p>कल के लक्ष्यों के लिए शुभकामनाएं।</p>
                <br>
                <p style="margin-bottom: 5px;">सादर,</p>
                <p style="margin-top: 0; font-weight: bold; color: #00d2ff;">AIR-01 RAS Automation System</p>
            </div>
            <div style="background: #f1f5f9; padding: 15px; text-align: center; font-size: 12px; color: #64748b;">
                यह एक स्वचालित रिमाइन्डर है। कृपया इसका उत्तर न दें।
            </div>
        </div>
    </body>
    </html>
    """
    
    msg.attach(MIMEText(html_content, 'html'))

    try:
        server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
        server.starttls()
        server.login(config['sender_email'], config['sender_password'])
        server.send_message(msg)
        server.quit()
        print("✅ Evening Reminder Sent Successfully!")
    except Exception as e:
        print(f"❌ Failed to send reminder: {e}")

if __name__ == "__main__":
    send_evening_reminder()
