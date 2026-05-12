import openpyxl
tracker_file = r"f:\AIR-01 RAS\Study_Automation_System\data\Master_Tracker_Live.xlsx"
try:
    wb = openpyxl.load_workbook(tracker_file, read_only=True)
    print(f"Sheets: {wb.sheetnames}")
except Exception as e:
    print(f"Error: {e}")
