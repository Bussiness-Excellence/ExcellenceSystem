@echo off
title Hierarchy Update
echo =======================================================
echo Replacing hierarchy in Supabase with hierarchy_export.xlsx
echo =======================================================
echo.
cd /d "%~dp0"
node update_hierarchy.js
echo.
echo =======================================================
echo Script finished. Press any key to close.
echo =======================================================
pause
