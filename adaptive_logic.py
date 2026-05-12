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
    
    col_map = {'section': 0, 'topic': 1, 'status': 13, 'comp_date': 14, 'next_action': 15}
    for i, h in enumerate(headers):
        if not h: continue
        if any(x in h for x in ['विषय', 'section', 'subject']): col_map['section'] = i
        if any(x in h for x in ['topic', 'अध्याय', 'पाठ', 'शीर्षक', 'विवरण']): col_map['topic'] = i
        if any(x in h for x in ['status', 'स्थिति']): col_map['status'] = i
        if any(x in h for x in ['completion', 'पूर्ण', 'तारीख']): col_map['comp_date'] = i
        if any(x in h for x in ['next action', 'अगला']): col_map['next_action'] = i
    
    print(f"DEBUG: Final Column Mapping: {col_map}")

    today = datetime.now().date()
    classes_list, revisions = [], []
    
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        if not any(row) or len(row) < 2: continue
        section = str(row[col_map['section']]).strip() if row[col_map['section']] else ""
        topic = str(row[col_map['topic']]).strip() if row[col_map['topic']] else ""
        status = str(row[col_map['status']]).strip().lower() if col_map['status'] < len(row) and row[col_map['status']] else ""
        comp_date = row[col_map['comp_date']] if col_map['comp_date'] < len(row) else None
        
        if status != "done" and len(classes_list) < 3: 
            classes_list.append({'subject': section, 'topic': topic})
            
        if status == "done" and isinstance(comp_date, datetime):
            c_date = comp_date.date()
            diff = (today - c_date).days
            label = ""
            if diff == 1: label = "1st Revision"
            elif diff == 7: label = "2nd Revision"
            elif diff == 30: label = "3rd Revision"
            if label: revisions.append({'subject': section, 'topic': f"{topic} ({label})"})

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
