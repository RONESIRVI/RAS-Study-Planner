import os
import json
import adaptive_logic
import image_generator
import excel_generator
import mailer
from datetime import datetime, timedelta, timezone

IST = timezone(timedelta(hours=5, minutes=30))

def run_automation():
    # Load config for time and weekend settings
    config_path = 'config.json'
    with open(config_path, 'r') as f:
        config = json.load(f)
    
    now = datetime.now(IST)
    current_day = now.strftime('%A')
    
    # Check if current day is weekend for logging schedule times
    is_current_weekend = current_day in ['Saturday', 'Sunday']
    if is_current_weekend:
        scheduled_time = config.get("weekend_morning_time", "04:00")
    else:
        scheduled_time = config.get("weekday_morning_time", "04:00")
    
    weekend_mode = config.get("weekend_mode", True)

    print(f"--- Starting Adaptive Study Automation at {now.strftime('%Y-%m-%d %H:%M:%S')} (IST) ---")
    print(f"Scheduled Time: {scheduled_time} | Weekend Mode: {'ON' if weekend_mode else 'OFF'}")
    
    try:
        tomorrow = now + timedelta(days=1)
        is_target_weekend = tomorrow.strftime('%A') in ['Saturday', 'Sunday']
        
        if weekend_mode and is_target_weekend:
            print(f"Target day is {tomorrow.strftime('%A')} (Weekend). Sending Weekly Review Plan.")
            custom_subject = f"📅 RAS Weekend Review & Plan - {now.strftime('%d %b %Y')}"
            extra_msg = "\n\nनोट: आज वीकेंड है! आज का दिन पिछले सप्ताह के विषयों के रिवीज़न और बैकलग (Backlog) क्लियर करने के लिए है।"
        else:
            custom_subject = None
            extra_msg = ""

        # 1. Fetch Adaptive Tasks and Revisions for tomorrow
        image_data, classes_list, progress_stats = adaptive_logic.get_adaptive_tasks(tomorrow)
        
        # Save tasks to JSON for Web Dashboard (Standard name)
        json_path = os.path.join("output", "todays_tasks.json")
        os.makedirs("output", exist_ok=True)
        with open(json_path, 'w', encoding='utf-8') as jf:
            json.dump(image_data, jf, ensure_ascii=False, indent=4)
            
        # Save tasks with Date-specific name to preserve history for Dashboard
        date_str = tomorrow.strftime('%Y-%m-%d')
        date_json_path = os.path.join("output", f"plan_{date_str}.json")
        with open(date_json_path, 'w', encoding='utf-8') as jf:
            json.dump(image_data, jf, ensure_ascii=False, indent=4)
            
        # Save progress stats to JSON for Web Dashboard
        progress_path = os.path.join("output", "progress.json")
        with open(progress_path, 'w', encoding='utf-8') as pf:
            json.dump(progress_stats, pf, ensure_ascii=False, indent=4)
        
        if not classes_list:
            print("INFO: No pending tasks found for today. (Check if everything is marked 'done')")
        
        # 2. Generate Premium Image with explicit target date
        print("Generating schedule image...")
        img_path = image_generator.create_pillar_schedule_image(image_data, target_date=tomorrow)
        print(f"DONE: Image generated: {img_path}")
        
        # 3. Generate PYQ Excel Files for each class (Disabled as per request)
        # print(f"Generating PYQ Excel files for {len(classes_list)} topics...")
        # pyq_files = []
        # for task in classes_list:
        #     pyq_file = excel_generator.generate_topic_excel(task['topic'])
        #     if pyq_file:
        #         pyq_files.append(pyq_file)
        
        # 4. Prepare Attachments
        attachments = [img_path]
        
        # 5. Send Email with explicit target date
        print(f"Preparing to send email with {len(attachments)} attachments...")
        mailer.send_email(attachments, extra_msg=extra_msg, custom_subject=custom_subject, target_date=tomorrow)
        
        print("SUCCESS: All tasks completed.")
        
    except Exception as e:
        print(f"ERROR: Error in automation flow: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    run_automation()
