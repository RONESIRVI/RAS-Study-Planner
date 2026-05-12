import openpyxl
import os
import glob
import requests
from datetime import datetime, timedelta, date

SHEET_ID = "1Zo81TfPcU09ErH7g-bj-4TksqceuBmiL"
DOWNLOAD_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"

def get_tracker_path():
    # Potential search directories
    search_dirs = [
        r"R:\Study_Automation_System",
        r"R:\Study_Automation_System\data",
        os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop"),
        os.path.join(os.path.expanduser("~"), "Desktop")
    ]
    
    found_files = []
    for d in search_dirs:
        if os.path.exists(d):
            # Find all xlsx files with 'Master Tracker' in name
            pattern = os.path.join(d, "*Master Tracker*.xlsx")
            found_files.extend(glob.glob(pattern))
            # Also check for 'Master_Tracker_Live'
            found_files.extend(glob.glob(os.path.join(d, "Master_Tracker_Live.xlsx")))

    if not found_files:
        return os.path.join("data", "Master_Tracker_Live.xlsx")
    
    # Sort by modification time (Newest first)
    found_files.sort(key=os.path.getmtime, reverse=True)
    print(f"DEBUG: Found {len(found_files)} potential trackers. Choosing newest: {found_files[0]}")
    return found_files[0]

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
        target_sheet = sheet_names[0] # Fallback to first sheet
        
    sheet = wb[target_sheet]
    print(f"DEBUG: Using sheet: {target_sheet}")
    
    header_row = 3
    all_headers = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    headers = [str(c).lower().strip() if c else "" for c in all_headers]
    print(f"DEBUG: Found headers: {headers}")
    
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

    today_day = today.weekday()  # 0=Mon, 5=Sat, 6=Sun
    is_weekend = today_day >= 5
    
    classes_list, revisions = [], []
    
    # Pre-clean headers for easier matching
    clean_headers = [str(h).lower().strip().replace('\n', ' ') for h in headers]
    
    # Find column indices dynamically and reliably
    col_map = {'section': 0, 'topic': 1, 'status': 13}
    for i, h in enumerate(clean_headers):
        if 'विषय' in h or 'section' in h: col_map['section'] = i
        if 'topic' in h: col_map['topic'] = i
        if 'status' in h: col_map['status'] = i
        if 'study date' in h: col_map['study_date'] = i
        if 'r1 date' in h: col_map['r1_date'] = i

    # Re-calculate indices safely
    topic_idx = 1
    section_idx = 0
    status_idx = 13
    
    for i, h in enumerate(headers):
        h_str = str(h).lower()
        if h_str == 'विषय' or h_str == 'section': section_idx = i
        if h_str == 'topic': topic_idx = i # Strict match for 'topic'
        if h_str == 'status': status_idx = i

    print(f"DEBUG: Reading File -> {os.path.abspath(tracker_file)}")
    
    count = 0
    # Sequential Picker: Pick first 2 pending tasks
    for row_idx, row in enumerate(sheet.iter_rows(min_row=4, max_row=200, values_only=True), 4):
        if not row: continue
        
        section = str(row[section_idx]).strip() if row[section_idx] is not None else ""
        topic = str(row[topic_idx]).strip() if row[topic_idx] is not None else ""
        
        raw_status = row[status_idx] if status_idx < len(row) else None
        status = str(raw_status).strip().lower() if raw_status is not None else ""
        
        # Super Safe Debug for first few rows
        if row_idx <= 6:
            print(f"DEBUG: Row {row_idx} | Topic: '{topic}' | Raw Status: '{raw_status}'")

        if topic and status != "done":
            if len(classes_list) < 2:
                print(f"DEBUG: ✅ PICKED Topic at Row {row_idx}: {topic}")
                classes_list.append({'subject': section, 'topic': topic})
                revisions.append({'subject': section, 'topic': f"{topic} (Same Day Rev)"})
            count += 1
            
        # 2. Spaced Repetition (ALWAYS RUNS for all subjects)
        for r_key, label in [('r1_date', 'R1'), ('r2_date', 'R2')]:
            rev_date = row[col_map[r_key]] if r_key in col_map and col_map[r_key] < len(row) else None
            if rev_date and isinstance(rev_date, datetime) and rev_date.date() == today:
                rev_done_key = f"{r_key.split('_')[0]}_done"
                rev_status = str(row[col_map[rev_done_key]]).strip().lower() if rev_done_key in col_map and col_map[rev_done_key] < len(row) else ""
                if rev_status != "done":
                    revisions.append({'subject': section, 'topic': f"{topic} ({label})"})
            
    print(f"DEBUG: Total classes added: {len(classes_list)}")

    # Format topic for image display (Keep separate)
    def get_c(idx):
        if idx < len(classes_list): return classes_list[idx]
        return {'subject': '[Complete]', 'topic': ''}

    image_data = []
    for idx, c in enumerate(classes_list):
        image_data.append({'sr': idx+1, 'task': f'CLASSES {idx+1}', 'subject': c['subject'], 'topic': c['topic']})
    
    image_data.append({'sr': len(image_data)+1, 'task': 'REVISION', 'revisions': revisions})
    image_data.append({'sr': len(image_data)+1, 'task': 'Analysis', 'subject': 'PYQ Review', 'topic': 'Previous Year Questions Analysis'})
    
    return image_data, classes_list

def get_weekly_roadmap():
    """Generates a list of the next 14 topics for a 7-day projection."""
    tracker_file = get_tracker_path()
    wb = openpyxl.load_workbook(tracker_file, data_only=True, read_only=True)
    sheet = wb['📋 Master Tracker']
    
    header_row = 3
    # Use the same logic as get_adaptive_tasks to find topics
    all_headers = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    headers = [str(c).lower().strip() if c else "" for c in all_headers]
    
    col_map = {'section': 0, 'topic': 1, 'status': 13}
    for i, h in enumerate(headers):
        if not h: continue
        if h == 'topic': col_map['topic'] = i
        if 'विषय' in h: col_map['section'] = i
        if 'status' in h: col_map['status'] = i

    roadmap = []
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        status = str(row[col_map['status']]).strip().lower() if col_map['status'] < len(row) and row[col_map['status']] else ""
        if status != "done" and len(roadmap) < 14:
            section = str(row[col_map['section']]).strip() if row[col_map['section']] else ""
            topic = str(row[col_map['topic']]).strip() if row[col_map['topic']] else ""
            if topic: roadmap.append(f"{section}: {topic}")
    
    return roadmap
