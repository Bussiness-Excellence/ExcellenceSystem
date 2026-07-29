#!/usr/bin/env python3
import os
import sys
import time
import argparse
import threading
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Add the parent directory to sys.path to import the main uploader
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from pulpoplus_upload_to_supabase import upload_workbook, _supabase_config, supabase_delete_where
    from pulpoplus_config import list_workbooks, wait_until_stable
except ImportError as e:
    print("Failed to import upload logic.")
    print("Error:", e)
    print("Make sure pulpoplus_upload_to_supabase.py and pulpoplus_config.py")
    print("are in this folder or its parent.")
    sys.exit(1)

def upload_folder(folder, period, batch):
    print(f"\n{'='*60}")
    print(f"Processing folder: {folder}")
    print(f"   Period: {period} | Batch: {batch}")

    # list_workbooks() filters out Excel's "~$" lock files, which the old
    # glob("*.xlsx") happily matched — and since a lock file is created the
    # instant somebody opens a workbook, it was always the newest by mtime,
    # so the uploader picked it and crashed on an unreadable file.
    files = list_workbooks(folder)
    if not files:
        print("   No .xlsx files found — skipping")
        return

    # Only process the SINGLE most recently modified file. Processing every
    # .xlsx in the folder used to combine old, already-uploaded exports with
    # the current one, producing duplicate coaching/visit rows any time an
    # old file was left in place instead of being removed.
    if len(files) > 1:
        print(f"   Found {len(files)} .xlsx files — using only the most recent one to avoid duplicating stale data:")
        for f in sorted(files, key=lambda p: p.stat().st_mtime, reverse=True):
            print(f"     {'-> ' if f == max(files, key=lambda p: p.stat().st_mtime) else '   '}{f.name}")
    latest = max(files, key=lambda p: p.stat().st_mtime)

    # A file that is still being written (or synced) reads as truncated.
    if not wait_until_stable(latest):
        print(f"   {latest.name} is still changing after 60s — skipping this round")
        return

    url, key = _supabase_config()

    file_batch = f"{batch}_{latest.stem}"
    print(f"   Uploading {latest.name} (Batch: {file_batch})...")
    try:
        upload_workbook(url, key, str(latest), period=period, batch=file_batch, append=True)
    except Exception as e:
        print(f"   Failed to upload {latest.name}: {e}")

    print(f"\n   Done processing {folder}!")

class FolderHandler(FileSystemEventHandler):
    def __init__(self, folder, period, batch):
        self.folder = folder
        self.period = period
        self.batch = batch
        self._pending = False
        self._restart = False
        self._lock = threading.Lock()

    def on_any_event(self, event):
        if event.is_directory:
            return
        name = os.path.basename(str(event.src_path))
        # Ignore Excel lock files and temp files, which fire constantly while
        # a workbook is simply open and would trigger endless re-uploads.
        if not name.lower().endswith(".xlsx") or name.startswith("~$") or name.startswith("."):
            return
        with self._lock:
            if self._pending:
                self._restart = True   # remember work arrived mid-run
                return
            self._pending = True

        def run():
            try:
                while True:
                    time.sleep(5)  # debounce burst writes
                    with self._lock:
                        self._restart = False
                    try:
                        upload_folder(self.folder, self.period, self.batch)
                    except Exception as exc:
                        print(f"   Upload failed: {exc}")
                    with self._lock:
                        if not self._restart:
                            self._pending = False
                            return
            except Exception as exc:
                print(f"   Watcher thread error: {exc}")
                with self._lock:
                    self._pending = False

        threading.Thread(target=run, daemon=True).start()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--last-month", default=r"e:\crm extractor\Periods\last_month", help="Folder for last month data")
    parser.add_argument("--recent", default=r"e:\crm extractor\Periods\recent", help="Folder for recent data")
    parser.add_argument("--once", action="store_true", help="Upload once and exit")
    args = parser.parse_args()

    folders = [
        (args.last_month, "Last Month", "last_month"),
        (args.recent, "Recent", "recent"),
    ]

    for folder, _, _ in folders:
        Path(folder).mkdir(parents=True, exist_ok=True)
        print(f"Folder ready: {folder}")

    for folder, period, batch in folders:
        upload_folder(folder, period, batch)

    if args.once:
        print("\nOne-time upload complete.")
        return

    print(f"\nWatching folders for new files... (Ctrl+C to stop)")
    observer = Observer()
    for folder, period, batch in folders:
        observer.schedule(FolderHandler(folder, period, batch), folder, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\nStopped.")
    observer.join()

if __name__ == "__main__":
    main()