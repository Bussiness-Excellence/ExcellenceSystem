@echo off
title Excellence CRM - One-Time Upload
echo =======================================================
echo  Uploading latest Excel file to Supabase...
echo =======================================================
echo.

cd /d "%~dp0"

python -c "
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath('.')))
from pathlib import Path

# Import upload logic directly (no watchdog needed)
sys.path.insert(0, r'e:\crm extractor')
from pulpoplus_upload_to_supabase import upload_workbook, _supabase_config

def upload_folder(folder, period, batch):
    files = list(Path(folder).glob('*.xlsx'))
    if not files:
        print(f'No .xlsx files found in {folder} - skipping')
        return
    latest = max(files, key=lambda p: p.stat().st_mtime)
    print(f'Found: {latest.name}')
    url, key = _supabase_config()
    file_batch = f'{batch}_{latest.stem}'
    print(f'Uploading with batch={file_batch}, period={period}...')
    upload_workbook(url, key, str(latest), period=period, batch=file_batch, append=True)
    print(f'Done!')

upload_folder(r'e:\crm extractor\Periods\recent', 'Recent', 'recent')
upload_folder(r'e:\crm extractor\Periods\last_month', 'Last Month', 'last_month')
"

echo.
echo =======================================================
echo  Upload complete!
echo =======================================================
pause
