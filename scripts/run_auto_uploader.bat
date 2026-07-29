@echo off
setlocal
title Excellence CRM Auto-Uploader
echo =======================================================
echo Starting Excellence CRM Excel Watcher ^& Uploader...
echo Watches the folders configured in .env:
echo   PERIOD_LAST_MONTH_DIR  (Last Month data)
echo   PERIOD_RECENT_DIR      (Recent 1-15 data)
echo =======================================================
echo.
cd /d "%~dp0"

if not exist ".env" (
    echo  ERROR: no .env file found in this folder.
    echo  Copy .env.example to .env and fill in your Supabase keys.
    pause
    exit /b 1
)

python pulpoplus_auto_upload.py %*
if errorlevel 1 (
    echo.
    echo  The uploader exited with an error - see above.
)
pause
