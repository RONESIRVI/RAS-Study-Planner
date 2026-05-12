@echo off
echo Pushing R: Drive changes to GitHub...
R:
cd \Study_Automation_System
git add .
git commit -m "Final Secure Setup: R-Drive, Weekly Roadmap, and Fixes"
git push
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ ERROR: Push failed! 
) else (
    echo.
    echo ✅ SUCCESS! System updated on GitHub.
)
pause
