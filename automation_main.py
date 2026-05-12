import os
import json
import adaptive_logic
import image_generator
import excel_generator
import mailer
from datetime import datetime

def run_automation():
    print(f"--- Starting Adaptive Study Automation at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")
    
    try:
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
            pyq_file = excel_generator.generate_pyq_excel(task['topic'])
            if pyq_file:
                pyq_files.append(pyq_file)
        
        # 4. Prepare Attachments
        attachments = [img_path] + pyq_files
        
        # 5. Send Email
        print(f"Preparing to send email with {len(attachments)} attachments...")
        mailer.send_email(attachments)
        
        print("SUCCESS: All tasks completed.")
        
    except Exception as e:
        print(f"ERROR: Error in automation flow: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    run_automation()
