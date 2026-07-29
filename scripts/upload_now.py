#!/usr/bin/env python3
"""
upload_now.py — one-time upload of the newest workbook in each period folder.

This replaces the Python that used to be embedded inside upload_now.bat.
That never ran: cmd.exe ends a command at the newline, so `python -c "` was an
unterminated argument and every following line was handed to the shell as its
own command.

Usage:
    python upload_now.py
    python upload_now.py --recent "D:\\data\\recent" --last-month "D:\\data\\last_month"
    python upload_now.py --replace     # clear each batch first instead of appending
"""

import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pulpoplus_config import env_path, list_workbooks, wait_until_stable  # noqa: E402
from pulpoplus_upload_to_supabase import upload_workbook, _supabase_config  # noqa: E402


def upload_folder(folder, period, batch, append=True):
    print(f"\n{'=' * 60}")
    print(f"Folder: {folder}")
    print(f"  Period: {period} | Batch prefix: {batch}")

    if not Path(folder).is_dir():
        print(f"  Folder does not exist — skipping")
        return False

    files = list_workbooks(folder)
    if not files:
        print("  No .xlsx files found — skipping")
        return False

    latest = max(files, key=lambda p: p.stat().st_mtime)
    if len(files) > 1:
        print(f"  {len(files)} workbooks present; using the most recent:")
        for f in sorted(files, key=lambda p: p.stat().st_mtime, reverse=True):
            print(f"    {'-> ' if f == latest else '   '}{f.name}")

    if not wait_until_stable(latest):
        print(f"  {latest.name} is still being written — skipping")
        return False

    url, key = _supabase_config()
    file_batch = f"{batch}_{latest.stem}"
    print(f"  Uploading {latest.name}  (batch={file_batch})")

    try:
        upload_workbook(url, key, str(latest), period=period,
                        batch=file_batch, append=append)
        return True
    except Exception as exc:
        print(f"  FAILED: {exc}")
        return False


def main():
    ap = argparse.ArgumentParser(description="One-time upload to Supabase.")
    ap.add_argument("--recent",
                    default=env_path("PERIOD_RECENT_DIR", r"E:\crm extractor\Periods\recent"))
    ap.add_argument("--last-month", dest="last_month",
                    default=env_path("PERIOD_LAST_MONTH_DIR", r"E:\crm extractor\Periods\last_month"))
    ap.add_argument("--replace", action="store_true",
                    help="Delete each batch's existing rows before inserting")
    args = ap.parse_args()

    append = not args.replace
    results = [
        upload_folder(args.recent, "Recent", "recent", append),
        upload_folder(args.last_month, "Last Month", "last_month", append),
    ]

    print("\n" + "=" * 60)
    if any(results):
        print("Upload complete.")
        return 0
    print("Nothing was uploaded — check the folder paths above.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
