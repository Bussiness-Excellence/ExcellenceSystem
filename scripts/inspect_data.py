#!/usr/bin/env python3
"""
inspect_data.py — read-only sanity check of the hierarchy export and the
matching Supabase tables.

FIXED:
  - Credentials come from .env, not from the source.
  - Uses the ANON key deliberately: this script only reads, so it should not
    be holding a key that can delete anything.
  - Excel path is configurable.
  - Reports HTTP failures clearly instead of dumping raw text.

Usage:
    python inspect_data.py
    python inspect_data.py "C:\\path\\to\\hierarchy_export_FIXED.xlsx"
"""

import sys
import json
from pathlib import Path

import pandas as pd
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pulpoplus_config import supabase_config, env_path  # noqa: E402


def show_table(url, headers, table, params, label):
    print(f"\nFetching {label} from Supabase...")
    try:
        res = requests.get(f"{url}/rest/v1/{table}", headers=headers, params=params, timeout=30)
    except requests.RequestException as exc:
        print(f"  Request failed: {exc}")
        return
    if res.status_code == 200:
        print(json.dumps(res.json(), indent=2, default=str))
    elif res.status_code in (401, 403):
        print(f"  HTTP {res.status_code}: the anon key cannot read `{table}`.")
        print("  That is expected if Row Level Security blocks anonymous reads.")
    else:
        print(f"  HTTP {res.status_code}: {res.text[:500]}")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    excel_path = args[0] if args else env_path(
        "HIERARCHY_FILE", r"C:\Users\Administrator\Downloads\hierarchy_export_FIXED.xlsx"
    )

    if Path(excel_path).is_file():
        try:
            df = pd.read_excel(excel_path)
            print(f"Excel: {excel_path}")
            print(f"  {len(df)} rows")
            print(f"  Columns: {df.columns.tolist()}")
            print("  First 3 rows:")
            for row in df.head(3).to_dict("records"):
                print(f"    {row}")
        except Exception as exc:
            print(f"Error reading Excel: {exc}")
    else:
        print(f"Excel file not found: {excel_path}")
        print("  (set HIERARCHY_FILE in .env or pass a path) — continuing with the DB check")

    # Read-only: anon key is the right level of access here.
    url, key = supabase_config(require_service_role=False)
    headers = {"apikey": key, "Authorization": f"Bearer {key}"}

    show_table(url, headers, "hierarchy", {"select": "*", "limit": "3"}, "hierarchy (first 3 rows)")
    show_table(url, headers, "teams", {"select": "*"}, "teams")


if __name__ == "__main__":
    main()
