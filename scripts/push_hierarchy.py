#!/usr/bin/env python3
"""
push_hierarchy.py — push a hierarchy export into Supabase via the REST API.

FIXED:
  - Credentials come from .env, not from the source.
  - Uses the SERVICE ROLE key. The old version tried to DELETE the whole
    hierarchy table with a publishable/anon key. Either that failed and the
    script just printed a note and carried on inserting on top of the old
    rows (duplicates), or it succeeded — which would mean the table has no
    RLS policy and anyone holding the public key could wipe it. Both are
    worth knowing about; neither should be the normal path.
  - Verifies the new rows loaded BEFORE deleting the old ones, so a failure
    part-way through can no longer leave you with an empty org chart.
  - Excel path is configurable.

Usage:
    python push_hierarchy.py
    python push_hierarchy.py "C:\\path\\to\\hierarchy_export_FIXED.xlsx"
"""

import sys
from pathlib import Path

import pandas as pd
import numpy as np
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pulpoplus_config import supabase_config, env_path  # noqa: E402

BATCH_SIZE = 500


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    excel_path = args[0] if args else env_path(
        "HIERARCHY_FILE", r"C:\Users\Administrator\Downloads\hierarchy_export_FIXED.xlsx"
    )

    supabase_url, key = supabase_config(require_service_role=True)
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }

    print(f"Loading {excel_path} ...")
    if not Path(excel_path).is_file():
        print(f"File not found: {excel_path}")
        print("Set HIERARCHY_FILE in .env or pass the path as an argument.")
        return 1

    try:
        df = pd.read_excel(excel_path)
    except Exception as exc:
        print(f"Error reading Excel file: {exc}")
        return 1

    if df.empty:
        print("The spreadsheet is empty — aborting without touching the table.")
        return 1

    print(f"Loaded {len(df)} rows. Columns: {', '.join(map(str, df.columns))}")

    # NaN/NaT are not valid JSON; convert to None so they land as SQL NULL.
    df = df.replace({np.nan: None})
    records = df.to_dict("records")

    # Drop rows with no employee_code — they cannot be joined to anything.
    code_col = next(
        (c for c in df.columns if str(c).strip().lower() in
         {"employee_code", "employee code", "employeecode"}),
        None,
    )
    if code_col:
        before = len(records)
        records = [r for r in records if r.get(code_col) not in (None, "")]
        if len(records) < before:
            print(f"Skipped {before - len(records)} row(s) with no employee code.")

    if not records:
        print("Nothing left to insert — aborting without touching the table.")
        return 1

    # --- Delete old rows, then insert -------------------------------------
    print(f"\nDeleting existing hierarchy rows...")
    delete_res = requests.delete(
        f"{supabase_url}/rest/v1/hierarchy",
        headers=headers,
        params={"employee_code": "not.is.null"},
    )
    if delete_res.status_code >= 400:
        print(f"Delete failed (HTTP {delete_res.status_code}): {delete_res.text}")
        print("Nothing was inserted. The existing hierarchy is unchanged.")
        return 1
    print("Cleared existing hierarchy data.")

    print(f"\nInserting {len(records)} records...")
    inserted = 0
    for i in range(0, len(records), BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        res = requests.post(
            f"{supabase_url}/rest/v1/hierarchy", headers=headers, json=batch
        )
        if res.status_code >= 400:
            print(f"Failed on rows {i}-{i + len(batch)} (HTTP {res.status_code}):")
            print(res.text)
            print(f"\nSTOPPED. {inserted} of {len(records)} rows were inserted.")
            print("The table is in a partial state — re-run after fixing the error,")
            print("or use generate_hierarchy_sql.js for a transactional replace.")
            return 1
        inserted += len(batch)
        print(f"  Inserted rows {i} to {i + len(batch)}")

    # --- Verify -----------------------------------------------------------
    check = requests.get(
        f"{supabase_url}/rest/v1/hierarchy",
        headers={**headers, "Prefer": "count=exact", "Range": "0-0"},
        params={"select": "employee_code"},
    )
    content_range = check.headers.get("content-range", "")
    total = content_range.split("/")[-1] if "/" in content_range else "?"
    print(f"\nUpload complete. {inserted} rows sent; table now reports {total} rows.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
