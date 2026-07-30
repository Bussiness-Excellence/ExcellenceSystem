#!/usr/bin/env python3
"""
remove_duplicates.py — Scans Supabase tables and removes any exact duplicate records.

Tables audited:
  - visits
  - summaries
  - coaching_days
  - specialty_classification
  - product_calls

Uses natural unique composite keys to identify duplicates, keeps the first occurrence,
and deletes extra duplicate rows via Supabase REST API using the service role key.
"""

import sys
from pathlib import Path
from collections import defaultdict
import requests

# Insert script directory into python path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from pulpoplus_config import supabase_config, clean_date


PAGE_SIZE = 1000


def _headers(key, prefer=None):
    h = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if prefer:
        h["Prefer"] = prefer
    return h


def fetch_all_rows(url, key, table, select="*"):
    """Fetch all rows from a Supabase table with pagination."""
    rows = []
    offset = 0
    while True:
        headers = _headers(key)
        headers["Range-Unit"] = "items"
        headers["Range"] = f"{offset}-{offset + PAGE_SIZE - 1}"
        res = requests.get(f"{url}/rest/v1/{table}", headers=headers, params={"select": select})
        if not res.ok:
            print(f"  ❌ Error fetching {table}: {res.status_code} {res.text}")
            return []
        page = res.json()
        rows.extend(page)
        if len(page) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
    return rows


def delete_ids(url, key, table, ids_to_delete):
    """Delete rows by ID in batches."""
    if not ids_to_delete:
        return 0
    deleted_count = 0
    batch_size = 100
    for i in range(0, len(ids_to_delete), batch_size):
        chunk = ids_to_delete[i:i + batch_size]
        # format in list for postgrest `in.(id1,id2,...)`
        id_str = ",".join(str(x) for x in chunk)
        res = requests.delete(
            f"{url}/rest/v1/{table}",
            headers=_headers(key),
            params={"id": f"in.({id_str})"}
        )
        if res.ok:
            deleted_count += len(chunk)
        else:
            print(f"  ⚠️ Deletion batch failed on {table}: {res.status_code} {res.text}")
    return deleted_count


def deduplicate_table(url, key, table, key_fn):
    """Scan table, find duplicate IDs based on key_fn, and delete duplicates."""
    print(f"\n🔍 Scanning `{table}` for duplicates...")
    rows = fetch_all_rows(url, key, table)
    print(f"   Fetched {len(rows)} total rows from `{table}`")

    if not rows:
        return

    grouped = defaultdict(list)
    for r in rows:
        natural_key = key_fn(r)
        grouped[natural_key].append(r)

    duplicate_ids = []
    for k, group in grouped.items():
        if len(group) > 1:
            # Keep group[0], mark group[1:] for deletion
            for dup in group[1:]:
                if "id" in dup and dup["id"] is not None:
                    duplicate_ids.append(dup["id"])

    if duplicate_ids:
        print(f"   ⚠️ Found {len(duplicate_ids)} duplicate row(s) out of {len(rows)} total.")
        print(f"   🧹 Cleaning up {len(duplicate_ids)} duplicate row(s)...")
        deleted = delete_ids(url, key, table, duplicate_ids)
        print(f"   ✅ Cleaned {deleted} duplicate row(s) from `{table}`.")
    else:
        print(f"   ✅ `{table}` is completely clean (0 duplicates found).")


def main():
    print("=" * 70)
    print("🧹 SUPABASE DUPLICATE REMOVAL ENGINE")
    print("=" * 70)

    try:
        url, key = supabase_config(require_service_role=True)
    except Exception as exc:
        print(f"Config error: {exc}")
        sys.exit(1)

    # 1. Visits table deduplication
    deduplicate_table(
        url, key, "visits",
        lambda r: (
            (r.get("user") or "").strip().lower(),
            clean_date(r.get("visit_date")),
            (r.get("visit_time") or "").strip(),
            (r.get("acc_id") or "").strip().lower(),
            (r.get("doctor_key") or "").strip().lower(),
            (r.get("upload_batch") or "").strip().lower()
        )
    )

    # 2. Summaries table deduplication
    deduplicate_table(
        url, key, "summaries",
        lambda r: (
            (str(r.get("employee_code") or "")).strip(),
            (r.get("user_name") or "").strip().lower(),
            (r.get("period") or "").strip().lower(),
            (r.get("upload_batch") or "").strip().lower()
        )
    )

    # 3. Coaching Days deduplication
    deduplicate_table(
        url, key, "coaching_days",
        lambda r: (
            (r.get("manager_name") or "").strip().lower(),
            (r.get("rep_name") or "").strip().lower(),
            clean_date(r.get("coaching_date")),
            (r.get("upload_batch") or "").strip().lower()
        )
    )

    # 4. Specialty Classification deduplication
    deduplicate_table(
        url, key, "specialty_classification",
        lambda r: (
            (str(r.get("employee_code") or "")).strip(),
            (r.get("user_name") or "").strip().lower(),
            (r.get("specialty") or "").strip().lower(),
            (r.get("classification") or "").strip().lower(),
            (r.get("shift") or "").strip().lower(),
            (r.get("period") or "").strip().lower(),
            (r.get("upload_batch") or "").strip().lower()
        )
    )

    # 5. Product Calls deduplication
    deduplicate_table(
        url, key, "product_calls",
        lambda r: (
            (str(r.get("employee_code") or "")).strip(),
            (r.get("user_name") or "").strip().lower(),
            (r.get("specialty") or "").strip().lower(),
            (r.get("product") or "").strip().lower(),
            (r.get("shift") or "").strip().lower(),
            (r.get("period") or "").strip().lower(),
            (r.get("upload_batch") or "").strip().lower()
        )
    )

    print("\n" + "=" * 70)
    print("✨ DEDUPLICATION PROCESS COMPLETE!")
    print("=" * 70)


if __name__ == "__main__":
    main()
