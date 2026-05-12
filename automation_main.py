import os
from datetime import datetime
import image_generator
import adaptive_logic
import excel_generator
import mailer
import json

CONFIG_FILE = 'config.json'
LOCAL_CONFIG = r'C:\Users\jlpms\OneDrive\Desktop\राजस्थान का इतिहास\Study_Automation\config.json'
ALT_CONFIG = os.path.join('..', 'राजस्थान का इतिहास', 'Study_Automation', 'config.json')

def run_adaptive_automation():
    print(f"Starting Adaptive Study Automation at {datetime.now()}")
    
    # 1. Get Adaptive Data
    try:
        tasks, assigned_topics = adaptive_logic.get_adaptive_tasks()
    except Exception as e:
        print(f"Error in adaptive logic: {e}")
        return

    # 2. Determine Mode
    is_weekend = datetime.now().weekday() >= 5
    mode = "Weekly" if is_weekend else "Daily"
    
    # 3. Generate Schedule Image
    attachments = []
    try:
        image_path = image_generator.create_pillar_schedule_image(tasks_data=tasks)
        attachments.append(image_path)
    except Exception as e:
        print(f"Error generating image: {e}")

    # 4. Generate PYQ Excel for today's classes
    for c_item in assigned_topics:
        try:
            topic_name = c_item.get('topic', '')
            excel_path = excel_generator.generate_topic_excel(topic_name)
            if excel_path: attachments.append(excel_path)
        except Exception as e:
            pass

    # 5. Send Email
    try:
        recipient = os.environ.get("RECIPIENT_EMAIL", "figuring.cse@gmail.com")
        if not recipient or recipient.strip() == "":
            recipient = "figuring.cse@gmail.com"
        
        # Override with config if available (Local only)
        if os.environ.get("GITHUB_ACTIONS") is None:
            config_to_use = None
            for path in [CONFIG_FILE, ALT_CONFIG, LOCAL_CONFIG]:
                if path and os.path.exists(path):
                    config_to_use = path
                    break
            if config_to_use:
                with open(config_to_use, 'r') as f:
                    config = json.load(f)
                    recipient = config.get("recipient_email", recipient)
        
        mailer.send_schedule_email(attachments, recipient)
        print(f"Success: Adaptive plan and Excel files sent to {recipient}.")
    except Exception as e:
        print(f"Error in automation: {e}")

if __name__ == "__main__":
    run_adaptive_automation()
