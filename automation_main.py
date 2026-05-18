import os
import json
import adaptive_logic
import image_generator
import excel_generator
import mailer
from datetime import datetime

def run_automation():
    # Load config for time and weekend settings
    config_path = 'config.json'
    with open(config_path, 'r') as f:
        config = json.load(f)
    
    now = datetime.now()
    current_day = now.strftime('%A')
    is_weekend = current_day in ['Saturday', 'Sunday']
    
    if is_weekend:
        scheduled_time = config.get("weekend_morning_time", "04:00")
    else:
        scheduled_time = config.get("weekday_morning_time", "04:00")
    
    weekend_mode = config.get("weekend_mode", True)

    print(f"--- Starting Adaptive Study Automation at {now.strftime('%Y-%m-%d %H:%M:%S')} ---")
    print(f"Scheduled Time: {scheduled_time} | Weekend Mode: {'ON' if weekend_mode else 'OFF'}")
    
    try:
        if weekend_mode and is_weekend:
            print(f"Today is {current_day} (Weekend). Sending Weekly Review Plan.")
            custom_subject = f"📅 RAS Weekend Review & Plan - {now.strftime('%d %b %Y')}"
            extra_msg = "\n\nनोट: आज वीकेंड है! आज का दिन पिछले सप्ताह के विषयों के रिवीज़न और बैकलग (Backlog) क्लियर करने के लिए है।"
        else:
            custom_subject = None
            extra_msg = ""
        # 1. Fetch Adaptive Tasks and Revisions
        image_data, classes_list = adaptive_logic.get_adaptive_tasks()
        
        if not classes_list:
            print("INFO: No pending tasks found for today. (Check if everything is marked 'done')")
        
        # 2. Generate Premium Image
        print("Generating schedule image...")
        img_path = image_generator.create_pillar_schedule_image(image_data)
        print(f"DONE: Image generated: {img_path}")
        
        # 3. Generate PYQ Excel Files for each class
        print(f"Generating PYQ Excel files for {len(classes_list)} topics...")
        pyq_files = []
        for task in classes_list:
            pyq_file = excel_generator.generate_topic_excel(task['topic'])
            if pyq_file:
                pyq_files.append(pyq_file)
        
        # 4. Prepare Attachments
        attachments = [img_path] + pyq_files
        
        # 5. Send Email
        print(f"Preparing to send email with {len(attachments)} attachments...")
        mailer.send_email(attachments, extra_msg=extra_msg, custom_subject=custom_subject)
        
        print("SUCCESS: All tasks completed.")
        
    except Exception as e:
        print(f"ERROR: Error in automation flow: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    run_automation()
