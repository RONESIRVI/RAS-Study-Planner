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
    
    msg = MIMEMultipart()
    msg['From'] = config['sender_email']
    msg['To'] = config['recipient_email']
    msg['Subject'] = f"📊 Daily Progress Report - {datetime.now().strftime('%d %b %Y')}"

    # Professional HTML Body
    html_content = f"""
    <html>
    <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 10px; overflow: hidden;">
            <div style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 30px; text-align: center; color: white;">
                <h1 style="margin: 0; font-size: 24px; letter-spacing: 1px;">Daily Milestone Review</h1>
            </div>
            <div style="padding: 40px; background: #ffffff;">
                <p>नमस्ते,</p>
                <p>आशा है कि आज का आपका अध्ययन (Study Session) फलदायी रहा होगा।</p>
                <p style="font-size: 16px;"><b>दिन का अंतिम कार्य:</b></p>
                <p>कृपया अपनी <b>Master Tracker Sheet</b> में आज के कार्यों का स्टेटस (Status) अपडेट कर दें। यदि आपने आज के टॉपिक्स पूरे कर लिए हैं, तो उन्हें <b>'done'</b> मार्क करना न भूलें ताकि कल का नया स्टडी प्लान समय पर जनरेट हो सके।</p>
                
                <div style="margin: 30px 0; padding: 20px; background: #f8fafc; border-left: 4px solid #00d2ff; border-radius: 4px;">
                    <i>"अनुशासन ही लक्ष्य और उपलब्धि के बीच का पुल है।"</i>
                </div>
                
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
