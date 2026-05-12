import openpyxl
import os
import glob
from datetime import datetime

BASE_DIR = r'C:\Users\jlpms\OneDrive\Desktop\राजस्थान का इतिहास'
TRACKER_DIR = glob.glob(os.path.join(BASE_DIR, "*Master Tracker*"))[0]
TRACKER_FILE = glob.glob(os.path.join(TRACKER_DIR, "*Master Tracker RAS.xlsx"))[0]

def get_tasks_for_today():
    wb = openpyxl.load_workbook(TRACKER_FILE, data_only=True)
    sheet = wb['📋 Master Tracker']
    
    today = datetime.now().date()
    tasks = []
    revisions = []
    
    # Iterate through rows (Skip header rows)
    for row in sheet.iter_rows(min_row=3, values_only=True):
        if not any(row): continue
        
        topic = str(row[1]) if row[1] else ""
        sub_topic = str(row[2]) if row[2] else ""
        target_date = row[13] # Col 14
        priority = str(row[12]) if row[12] else ""
        
        # Check if target date is today
        if isinstance(target_date, datetime):
            if target_date.date() == today:
                tasks.append(f"{topic}: {sub_topic}")
        
        # Check for revision (e.g., Confidence < 3 or specific revision column)
        # For now, let's just pick High Priority items as tasks
        if priority.upper() == 'H' and f"{topic}: {sub_topic}" not in tasks:
            revisions.append(f"{topic}: {sub_topic}")

    # Format for Image Generator
    image_data = [
        {'sr': 1, 'task': 'CLASSES 1', 'topic': tasks[0] if len(tasks) > 0 else "[Leave Blank]"},
        {'sr': 2, 'task': 'CLASSES 2', 'topic': tasks[1] if len(tasks) > 1 else "[Leave Blank]"},
        {'sr': 3, 'task': 'REVISION', 'topic': f"Topics: {', '.join(revisions[:2])}" if revisions else "[No Revision Pending]"},
        {'sr': 4, 'task': 'Analysis', 'topic': '1. PYQ test\n2. MCQ Test'}
    ]
    
    # If Saturday/Sunday, add Weekly Test
    if datetime.now().weekday() >= 5: # 5=Sat, 6=Sun
        image_data.append({'sr': 5, 'task': 'WEEKLY TEST', 'topic': 'Comprehensive Mock Test\nSectional Evaluation'})
        
    return image_data

if __name__ == "__main__":
    data = get_tasks_for_today()
    print(data)
