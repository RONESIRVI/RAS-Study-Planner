@echo off
echo Running Study Automation Advanced Debug...
R:
cd \Study_Automation_System
:: Run and capture both output and errors to a file
python automation_main.py > C:\Users\jlpms\OneDrive\Desktop\log_report.txt 2>&1
echo.
echo Process finished. Checking logs...
pause
