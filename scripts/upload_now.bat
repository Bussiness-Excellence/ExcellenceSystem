@echo off
setlocal
title Excellence CRM - One-Time Upload
echo =======================================================
echo  Uploading latest Excel file to Supabase...
echo =======================================================
echo.

cd /d "%~dp0"

REM The Python that used to be inline here could never run: cmd.exe ends a
REM command at the newline, so the opening quote of python -c " was never
REM closed and every following line was executed as a separate command.
REM It now lives in a real script file.
python upload_now.py %*
if errorlevel 1 (
    echo.
    echo  Upload FAILED - see the messages above.
    pause
    exit /b 1
)

echo.
echo =======================================================
echo  Upload complete!
echo =======================================================
pause
