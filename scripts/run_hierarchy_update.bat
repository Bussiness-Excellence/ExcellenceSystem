@echo off
setlocal
title Hierarchy Update
echo =======================================================
echo Replacing hierarchy in Supabase from the hierarchy export
echo =======================================================
echo.
cd /d "%~dp0"

if not exist ".env" (
    echo  ERROR: no .env file found in this folder.
    echo  Copy .env.example to .env and fill in your Supabase keys.
    pause
    exit /b 1
)
if not exist "node_modules" (
    echo  WARNING: node_modules not found. Run "npm install" first.
    echo.
)

node update_hierarchy.js %*
if errorlevel 1 (
    echo.
    echo  Hierarchy update FAILED - the table was left untouched.
    pause
    exit /b 1
)

echo.
echo =======================================================
echo Script finished. Press any key to close.
echo =======================================================
pause
