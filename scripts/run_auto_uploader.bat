@echo off
title Excellence CRM Auto-Uploader
echo =======================================================
echo Starting Excellence CRM Excel Watcher & Uploader...
echo Watches:
echo   - E:\periods\last_month\ (Last Month data)
echo   - E:\periods\recent\     (Recent 1-15 data)
echo =======================================================
echo.
cd /d "%~dp0"
python pulpoplus_auto_upload.py
pause
