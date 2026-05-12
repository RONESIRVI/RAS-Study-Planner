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
    
    print(f"DEBUG: Cloud Sync -> Downloading latest from Google Drive...")
    try:
        r = requests.get(DOWNLOAD_URL, timeout=15)
        if r.status_code == 200:
            with open(target_path, "wb") as f:
                f.write(r.content)
            print("DEBUG: ✅ Download Success!")
            return target_path
    except Exception as e:
        print(f"DEBUG: ⚠️ Download Failed ({e}). Using local fallback...")

    # Local search if download fails
    search_dirs = [r"R:\Study_Automation_System\data", "data", "."]
    for d in search_dirs:
        p = os.path.join(d, "Master_Tracker_Live.xlsx")
        if os.path.exists(p): return p
    
    return target_path

def get_adaptive_tasks():
    today = datetime.now().date()
    tracker_file = get_tracker_path()
    wb = openpyxl.load_workbook(tracker_file, data_only=True, read_only=True)
    
    # Try multiple sheet names to be safe
    sheet_names = [sheet.title for sheet in wb.worksheets]
    target_sheet = None
    for name in ['Master Tracker', '📋 Master Tracker', 'Master_Tracker']:
        if name in sheet_names:
            target_sheet = name
            break
    
    if not target_sheet:
        target_sheet = sheet_names[0]
        
    sheet = wb[target_sheet]
    print(f"DEBUG: Using sheet: {target_sheet}")
    
    header_row = 3
    all_headers = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    headers = [str(c).lower().strip() if c else "" for c in all_headers]
    
    # Pre-clean headers for easier matching
    clean_headers = [str(h).lower().strip().replace('\n', ' ') for h in headers]
    
    # Find column indices dynamically
    col_map = {'section': 0, 'topic': 1, 'status': 13, 'r1_date': 3, 'r2_date': 5, 'r1_done': 4, 'r2_done': 6}
    for i, h in enumerate(clean_headers):
        if 'विषय' in h or h == 'section': col_map['section'] = i
        if h == 'topic': col_map['topic'] = i
        if h == 'status': col_map['status'] = i
        if 'r1 date' in h: col_map['r1_date'] = i
        if 'r2 date' in h: col_map['r2_date'] = i
        if 'r1 done' in h: col_map['r1_done'] = i
        if 'r2 done' in h: col_map['r2_done'] = i

    section_idx = col_map['section']
    topic_idx = col_map['topic']
    status_idx = col_map['status']

    print(f"DEBUG: Reading File -> {os.path.abspath(tracker_file)}")
    print("-" * 50)
    print(f"MASTER DEBUG: Checking first 20 rows...")
    
    classes_list, revisions = [], []
    
    # Process rows 4 to 100
    for row_idx, row in enumerate(sheet.iter_rows(min_row=4, max_row=100, values_only=True), 4):
        if not row: continue
        
        section = str(row[section_idx]).strip() if section_idx < len(row) and row[section_idx] is not None else ""
        topic = str(row[topic_idx]).strip() if topic_idx < len(row) and row[topic_idx] is not None else ""
        raw_status = row[status_idx] if status_idx < len(row) else None
        status = str(raw_status).strip().lower() if raw_status is not None else ""
        
        if not topic: continue
            
        # 1. Classes Logic (Pick first 2 pending)
        if status != "done" and len(classes_list) < 2:
            print(f"Row {row_idx}: Topic='{topic[:20]}...' -> ✅ PICKED!")
            classes_list.append({'subject': section, 'topic': topic})
            revisions.append({'subject': section, 'topic': f"{topic} (Same Day Rev)"})
        
        # 2. Spaced Repetition (R1, R2)
        for r_key, label in [('r1_date', 'R1'), ('r2_date', 'R2')]:
            if r_key in col_map:
                rev_val = row[col_map[r_key]] if col_map[r_key] < len(row) else None
                if rev_val and isinstance(rev_val, datetime) and rev_val.date() == today:
                    done_key = f"{r_key.split('_')[0]}_done"
                    done_status = str(row[col_map[done_key]]).strip().lower() if done_key in col_map and col_map[done_key] < len(row) else ""
                    if done_status != "done":
                        revisions.append({'subject': section, 'topic': f"{topic} ({label})"})

    print("-" * 50)
    
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
