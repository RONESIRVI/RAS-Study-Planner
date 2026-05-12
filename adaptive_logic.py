import openpyxl
import os
import glob
import requests
from datetime import datetime, timedelta, date

SHEET_ID = "1Zo81TfPcU09ErH7g-bj-4TksqceuBmiL"
DOWNLOAD_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"

def get_tracker_path():
    target_path = os.path.join("data", "Master_Tracker_Live.xlsx")
    os.makedirs("data", exist_ok=True)
    try:
        r = requests.get(DOWNLOAD_URL, timeout=15)
        if r.status_code == 200:
            with open(target_path, "wb") as f:
                f.write(r.content)
            return target_path
    except:
        pass
    return target_path

def get_adaptive_tasks():
    today = datetime.now().date()
    tracker_file = get_tracker_path()
    wb = openpyxl.load_workbook(tracker_file, data_only=True, read_only=True)
    sheet_names = wb.sheetnames
    
    # 1. Get Main Tasks from Master Tracker
    main_sheet_name = next((s for s in ['📋 Master Tracker', 'Master Tracker'] if s in sheet_names), sheet_names[0])
    sheet = wb[main_sheet_name]
    
    header_row = 3
    all_headers = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    headers = [str(c).lower().strip() if c else "" for c in all_headers]
    
    col_map = {'section': 0, 'topic': 1, 'status': 13}
    for i, h in enumerate(headers):
        if 'विषय' in h: col_map['section'] = i
        if 'topic' in h: col_map['topic'] = i
        if 'status' in h: col_map['status'] = i

    classes_list, revisions = [], []
    
    # Pick next 2 classes
    for row in sheet.iter_rows(min_row=4, max_row=200, values_only=True):
        topic = str(row[col_map['topic']]).strip() if col_map['topic'] < len(row) and row[col_map['topic']] else ""
        status = str(row[col_map['status']]).strip().lower() if col_map['status'] < len(row) and row[col_map['status']] else ""
        section = str(row[col_map['section']]).strip() if col_map['section'] < len(row) and row[col_map['section']] else ""
        
        if topic and status != "done" and len(classes_list) < 2:
            classes_list.append({'subject': section, 'topic': topic})
            revisions.append({'subject': section, 'topic': f"{topic} (Same Day Rev)"})

    # 2. Get Revisions from "Revision Planner" Sheet
    if 'Revision Planner' in sheet_names:
        rev_sheet = wb['Revision Planner']
        # Search every cell for today's date
        for row in rev_sheet.iter_rows(min_row=2, values_only=True):
            if not any(row): continue
            
            sub = str(row[0]).strip() if row[0] else ""
            top = str(row[1]).strip() if row[1] else ""
            
            # Check all columns starting from 3rd for today's date
            for cell_val in row[2:]:
                if isinstance(cell_val, datetime) and cell_val.date() == today:
                    revisions.append({'subject': sub, 'topic': f"{top} (Scheduled)"})
                    break

    # Prepare Image Data
    image_data = []
    for idx, c in enumerate(classes_list):
        image_data.append({'sr': idx+1, 'task': f'CLASSES {idx+1}', 'subject': c['subject'], 'topic': c['topic']})
    
    image_data.append({'sr': len(image_data)+1, 'task': 'REVISION', 'revisions': revisions})
    image_data.append({'sr': len(image_data)+1, 'task': 'Analysis', 'subject': 'PYQ Review', 'topic': 'Previous Year Questions Analysis'})
    
    return image_data, classes_list

def get_weekly_roadmap():
    tracker_file = get_tracker_path()
    wb = openpyxl.load_workbook(tracker_file, data_only=True, read_only=True)
    sheet = wb.worksheets[0]
    roadmap = []
    for row in sheet.iter_rows(min_row=4, max_row=50, values_only=True):
        if row[1] and len(roadmap) < 14:
            roadmap.append(f"{row[0]}: {row[1]}")
    return roadmap
