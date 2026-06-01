import smtplib
import json
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import adaptive_logic

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
    today = now.date()
    current_day = now.strftime('%A')
    is_weekend = current_day in ['Saturday', 'Sunday']
    
    if is_weekend:
        scheduled_time = config.get("weekend_evening_time", "13:00")
    else:
        scheduled_time = config.get("weekday_evening_time", "13:00")
        
    weekend_mode = config.get("weekend_mode", True)

    print(f"--- Starting Evening Reminder at {now.strftime('%Y-%m-%d %H:%M:%S')} ---")
    print(f"Scheduled Time: {scheduled_time} | Weekend Mode: {'ON' if weekend_mode else 'OFF'}")

    # Fetch today's tasks to display in the email
    try:
        image_data, _ = adaptive_logic.get_adaptive_tasks(today)
    except Exception as e:
        print(f"Error fetching today's tasks: {e}")
        image_data = []

    classes = []
    revisions = []
    pyq_tests = []
    mock_tests = []

    for task in image_data:
        task_type = task.get('task', '')
        if 'CLASSES' in task_type:
            sub = task.get('subject', '').strip()
            top = task.get('topic', '').strip()
            if sub and '[' not in sub:
                classes.append(f"<li><b>{sub}</b>: {top}</li>")
                pyq_tests.append(f"<li><b>{sub}</b>: {top}</li>")
        elif task_type == 'REVISION':
            revs = task.get('revisions', [])
            for r in revs:
                sub_r = r['subject'].strip()
                top_r = r['topic'].strip()
                revisions.append(f"<li><b>{sub_r}</b>: {top_r}</li>")
                if "Same Day Rev" not in r.get('topic', ''):
                    mock_tests.append(f"<li><b>{sub_r}</b>: {top_r}</li>")

    classes_list_html = "".join(classes) if classes else "<li><i>कोई क्लास शेड्यूल नहीं थी</i></li>"
    revisions_list_html = "".join(revisions) if revisions else "<li><i>कोई रिवीज़न शेड्यूल नहीं था</i></li>"
    pyq_list_html = "".join(pyq_tests) if pyq_tests else "<li><i>कोई PYQ टेस्ट नहीं था</i></li>"
    mock_list_html = "".join(mock_tests) if mock_tests else "<li><i>कोई मॉक टेस्ट नहीं था</i></li>"

    msg = MIMEMultipart()
    msg['From'] = f"RAS Mentorship Reminder <{config['sender_email']}>"
    msg['To'] = config['recipient_email']
    
    if weekend_mode and is_weekend:
        msg['Subject'] = f"🕯️ Weekend Reflection - {now.strftime('%d %b %Y')}"
        title = "Weekend Reflection"
        message_body = f"""
        <p>नमस्ते,</p>
        <p>आज वीकेंड है, आशा है कि आपने अपनी पढ़ाई का रिवीज़न (Revision) अच्छे से किया होगा।</p>
        <p><b>वीकेंड टास्क:</b></p>
        <p>1. पिछले सप्ताह के छूटे हुए टॉपिक्स (Backlogs) को चेक करें।<br>
        2. अगले सप्ताह के लिए अपनी ऊर्जा संचित करें।<br>
        3. अपनी <b>Master Tracker Sheet</b> में आज के कार्यों का स्टेटस (उन्हें <b>'done'</b> या <b>'ok'</b> मार्क करना न भूलें) अपडेट कर दें।</p>
        
        <p><b>आज के भेजे गए कार्य (जिन्हें मास्टर शीट में OK/DONE मार्क करना है):</b></p>
        <ul style="padding-left: 20px;">
            <li style="margin-bottom: 10px;"><b>CLASSES (Daily Focus):</b>
                <ul style="padding-left: 15px; margin-top: 5px;">{classes_list_html}</ul>
            </li>
            <li style="margin-bottom: 10px;"><b>REVISION (Spaced Repetition):</b>
                <ul style="padding-left: 15px; margin-top: 5px;">{revisions_list_html}</ul>
            </li>
            <li style="margin-bottom: 10px;"><b>PYQ TEST (Assessment):</b>
                <ul style="padding-left: 15px; margin-top: 5px;">{pyq_list_html}</ul>
            </li>
            <li style="margin-bottom: 10px;"><b>MOCK TEST (Strengthen):</b>
                <ul style="padding-left: 15px; margin-top: 5px;">{mock_list_html}</ul>
            </li>
        </ul>

        <div style="margin: 30px 0; padding: 20px; background: #f8fafc; border-left: 4px solid #f59e0b; border-radius: 4px;">
            <i>"तैयारी करने में विफल रहने का मतलब है विफल होने की तैयारी करना।"</i>
        </div>
        """
    else:
        msg['Subject'] = f"📊 Daily Progress Report - {now.strftime('%d %b %Y')}"
        title = "Daily Milestone Review"
        message_body = f"""
        <p>नमस्ते,</p>
        <p>आशा है कि आज का आपका अध्ययन (Study Session) फलदायी रहा होगा।</p>
        <p>कल जो प्लान आज के लिए भेजा गया था, उसके अनुसार आज के कार्य निम्नलिखित हैं। कृपया इन्हें पूरा करने के बाद अपनी <b>Master Tracker Sheet</b> में स्टेटस को <b>'done'</b> या <b>'ok'</b> मार्क करना न भूलें:</p>
        
        <p><b>आज के भेजे गए कार्य (जिन्हें मास्टर शीट में OK/DONE मार्क करना है):</b></p>
        <ul style="padding-left: 20px;">
            <li style="margin-bottom: 10px;"><b>CLASSES (Daily Focus):</b>
                <ul style="padding-left: 15px; margin-top: 5px;">{classes_list_html}</ul>
            </li>
            <li style="margin-bottom: 10px;"><b>REVISION (Spaced Repetition):</b>
                <ul style="padding-left: 15px; margin-top: 5px;">{revisions_list_html}</ul>
            </li>
            <li style="margin-bottom: 10px;"><b>PYQ TEST (Assessment):</b>
                <ul style="padding-left: 15px; margin-top: 5px;">{pyq_list_html}</ul>
            </li>
            <li style="margin-bottom: 10px;"><b>MOCK TEST (Strengthen):</b>
                <ul style="padding-left: 15px; margin-top: 5px;">{mock_list_html}</ul>
            </li>
        </ul>

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
        print("[SUCCESS] Evening Reminder Sent Successfully!")
    except Exception as e:
        print(f"[ERROR] Failed to send reminder: {e}")

if __name__ == "__main__":
    send_evening_reminder()
