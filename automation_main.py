import os
import sys
import json
from datetime import datetime

# Fix for Windows Unicode issues
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

import image_generator
import adaptive_logic
import excel_generator
import mailer

# Configuration and Data Paths
POSSIBLE_CONFIGS = [
    'config.json',
    r'R:\Study_Automation_System\config.json',
    r'C:\Users\jlpms\OneDrive\Desktop\राजस्थान का इतिहास\Study_Automation\config.json'
]

search_paths = [
    os.path.join("data", "Master_Tracker_Live.xlsx"),
    "Master_Tracker_Live.xlsx",
    r'R:\Study_Automation_System\data\Master_Tracker_Live.xlsx',
    r'C:\Users\jlpms\OneDrive\Desktop\राजस्थान का इतिहास\📋 Master Tracker\Master Tracker RAS.xlsx'
]

def run_adaptive_automation():
    print(f"\n--- Starting Adaptive Study Automation at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")
    
    # 1. Get Adaptive Data
    try:
        tasks, assigned_topics = adaptive_logic.get_adaptive_tasks()
        print(f"✅ Data fetched: {len(assigned_topics)} classes and {len(tasks[2]['revisions'])} revisions found.")
    except Exception as e:
        print(f"❌ Error in adaptive logic: {e}")
        return

    # 2. Determine Mode
    is_weekend = datetime.now().weekday() >= 5
    mode = "Weekly" if is_weekend else "Daily"
    print(f"Mode: {mode}")
    
    # 3. Generate Schedule Image
    attachments = []
    try:
        print("Generating schedule image...")
        image_path = image_generator.create_pillar_schedule_image(tasks_data=tasks)
        if image_path and os.path.exists(image_path):
            attachments.append(image_path)
            print(f"✅ Image generated: {image_path}")
        else:
            print("⚠️ Image generation returned no path or file missing.")
    except Exception as e:
        print(f"❌ Error generating image: {e}")

    # 4. Generate PYQ Excel for today's classes
    print(f"Generating PYQ Excel files for {len(assigned_topics)} topics...")
    for c_item in assigned_topics:
        try:
            topic_name = c_item.get('topic', '')
            if not topic_name: continue
            excel_path = excel_generator.generate_topic_excel(topic_name)
            if excel_path and os.path.exists(excel_path):
                attachments.append(excel_path)
                print(f"✅ Excel generated for: {topic_name}")
            else:
                print(f"ℹ️ No PYQ found for topic: {topic_name}")
        except Exception as e:
            print(f"⚠️ Error generating Excel for {c_item.get('topic')}: {e}")

    # 5. Send Email
    try:
        recipient = os.environ.get("RECIPIENT_EMAIL", "figuring.cse@gmail.com")
        
        # Priority 1: Check Local config (R:, C: or current dir)
        config_to_use = None
        for path in POSSIBLE_CONFIGS:
            if path and os.path.exists(path):
                config_to_use = path
                break
        
        if config_to_use:
            print(f"Using config from: {config_to_use}")
            with open(config_to_use, 'r') as f:
                config = json.load(f)
                recipient = config.get("recipient_email", recipient)
        else:
            print("⚠️ No config.json found. Using environment variables or defaults.")

        print(f"Preparing to send email to {recipient} with {len(attachments)} attachments...")
        
        # Weekly Roadmap Logic (Trigger on Sundays)
        weekly_msg = ""
        if datetime.now().weekday() == 6: # Sunday
            roadmap = adaptive_logic.get_weekly_roadmap()
            weekly_msg = "\n\n🚀 --- UPCOMING WEEKLY ROADMAP ---\n"
            for i, item in enumerate(roadmap, 1):
                day_num = (i + 1) // 2
                if i % 2 != 0: weekly_msg += f"\nDay {day_num}:\n"
                weekly_msg += f"  - {item}\n"
        
        mailer.send_schedule_email(attachments, recipient, extra_msg=weekly_msg)
        print(f"🎉 Success: All tasks completed.")
    except Exception as e:
        print(f"❌ Critical error in automation pipeline: {e}")

if __name__ == "__main__":
    run_adaptive_automation()
