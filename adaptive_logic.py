import openpyxl
import os
import glob
import requests
from datetime import datetime, timedelta

SHEET_ID = "1Zo81TfPcU09ErH7g-bj-4TksqceuBmiL"
DOWNLOAD_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"

def get_tracker_path():
    target_path = os.path.join("data", "Master_Tracker_Live.xlsx")
    
    # 1. If on Cloud (GitHub)
    if os.environ.get("GITHUB_ACTIONS"):
        if os.path.exists(target_path): return target_path
        print("Cloud Mode: Downloading latest tracker...")
        try:
            os.makedirs("data", exist_ok=True)
            r = requests.get(DOWNLOAD_URL)
            with open(target_path, "wb") as f:
                f.write(r.content)
            return target_path
        except:
            return glob.glob(os.path.join("**", "*Master Tracker RAS.xlsx"), recursive=True)[0]
    
    # 2. If on Local PC, search multiple locations
    search_paths = [
        os.path.join("data", "Master_Tracker_Live.xlsx"),
        "Master_Tracker_Live.xlsx",
        os.path.join("..", "राजस्थान का इतिहास", "📋 Master Tracker", "Master Tracker RAS.xlsx"),
        r'C:\Users\jlpms\OneDrive\Desktop\राजस्थान का इतिहास\📋 Master Tracker\Master Tracker RAS.xlsx'
    ]
    for p in search_paths:
        if os.path.exists(p): return p
        
    # Fallback search
    found = glob.glob(os.path.join("**", "*Master Tracker RAS.xlsx"), recursive=True)
    return found[0] if found else target_path

def get_adaptive_tasks():
    tracker_file = get_tracker_path()
    wb = openpyxl.load_workbook(tracker_file, data_only=True, read_only=True)
    sheet = wb['📋 Master Tracker']
    
    header_row = 3
    all_headers = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    headers = [str(c).lower().strip() if c else "" for c in all_headers]
    
    print(f"DEBUG: Found headers in Excel: {headers}")
    
    col_map = {'section': 0, 'topic': 1, 'study_date': 2, 'r1_date': 3, 'r1_done': 4, 'r2_date': 5, 'r2_done': 6, 'status': 12, 'comp_date': 13}
    for i, h in enumerate(headers):
        if not h: continue
        h_str = h.lower()
        if h_str == 'topic': col_map['topic'] = i; continue
        if h_str == 'विषय': col_map['section'] = i; continue
        if 'study date' in h_str: col_map['study_date'] = i
        if 'r1 date' in h_str: col_map['r1_date'] = i
        if 'r1 done' in h_str: col_map['r1_done'] = i
        if 'status' in h_str: col_map['status'] = i
        if 'completion' in h_str or 't_date' in h_str: col_map['comp_date'] = i
    
    print(f"DEBUG: Final Column Mapping: {col_map}")

    today = datetime.now().date()
    classes_list, revisions = [], []
    
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(row) or len(row) < 2: continue
        section = str(row[col_map['section']]).strip() if row[col_map['section']] else ""
        topic = str(row[col_map['topic']]).strip() if row[col_map['topic']] else ""
        status = str(row[col_map['status']]).strip().lower() if col_map['status'] < len(row) and row[col_map['status']] else ""
        
        # 1. Classes Logic (Status not done)
        if status != "done" and len(classes_list) < 3: 
            classes_list.append({'subject': section, 'topic': topic})
            
        # 2. Revision Logic (Check specific dates)
        for r_key, label in [('r1_date', 'R1'), ('r2_date', 'R2')]:
            r_val = row[col_map.get(r_key)] if col_map.get(r_key) and col_map[r_key] < len(row) else None
            if isinstance(r_val, datetime) and r_val.date() == today:
                revisions.append({'subject': section, 'topic': f"{topic} ({label})"})

    # Format topic for image display (Keep separate)
    def get_c(idx):
        if idx < len(classes_list): return classes_list[idx]
        return {'subject': '[Complete]', 'topic': ''}

    image_data = [
        {'sr': 1, 'task': 'CLASSES 1', 'subject': get_c(0)['subject'], 'topic': get_c(0)['topic']},
        {'sr': 2, 'task': 'CLASSES 2', 'subject': get_c(1)['subject'], 'topic': get_c(1)['topic']},
        {'sr': 3, 'task': 'REVISION', 'revisions': revisions},
        {'sr': 4, 'task': 'Analysis', 'subject': 'PYQ Review', 'topic': 'Previous Year Questions Analysis'}
    ]
    return image_data, classes_list
