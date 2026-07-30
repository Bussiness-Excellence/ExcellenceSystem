"""
pulpoplus_upload_to_supabase.py

Pushes a rebuilt PulpoPlus summary workbook (produced by
pulpoplus_rebuild_summary.py / pulpoplus_extract_visits.py) into Supabase,
and/or syncs the team-structure hierarchy file into the `teams` and
`hierarchy` tables.

WHAT THIS DOES NOT DO: it does not create any Supabase Auth (login)
accounts, and does not touch the `app_users` table. That's a separate,
deliberate decision (email format, password scheme, etc.) — this script
only pushes data, using `employee_code` as a plain text join key, exactly
like your existing `visits`/`coaching_days` tables already do.

SETUP (one-time):
    pip install requests pandas openpyxl --break-system-packages

    Copy .env.example to .env and fill in:
        SUPABASE_URL               your project URL
        SUPABASE_SERVICE_ROLE_KEY  the "service_role" secret from
                                   Project Settings -> API (NOT the anon
                                   key -- the service role key is required
                                   to write through Row Level Security).

    Credentials are read from .env by pulpoplus_config. They are never
    stored in this file. If a key ever appears in source, rotate it in the
    Supabase dashboard immediately -- anything committed to git is public.

USAGE:
    # Sync the org structure (run whenever the team structure changes)
    python pulpoplus_upload_to_supabase.py --hierarchy "hierarchy/Team Structure.xlsx"

    # Upload a rebuilt month's data (run after every rebuild)
    python pulpoplus_upload_to_supabase.py --workbook "June 2026 - Platinum 1.xlsx"

    # Both in one go
    python pulpoplus_upload_to_supabase.py --hierarchy "hierarchy/Team Structure.xlsx" --workbook "June 2026 - Platinum 1.xlsx"

Re-running an upload for the same period REPLACES that period's rows
(matched by an `upload_batch` value) rather than duplicating them, so it's
always safe to re-run after fixing something.
"""

import os
import re
import sys
import argparse
from collections import Counter
from datetime import datetime

import pandas as pd
import requests

import pulpoplus_hierarchy as hier
from pulpoplus_config import supabase_config, clean_date, clean_time

CHUNK_SIZE = 500  # rows per HTTP request, to stay well under payload limits


# ── Supabase REST helpers ──────────────────────────────────────────────────

def _supabase_config():
    """Credentials from the environment / .env — never from source."""
    return supabase_config(require_service_role=True)


def _headers(key, prefer=None):
    h = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if prefer:
        h["Prefer"] = prefer
    return h


def supabase_delete_where(url, key, table, column, value):
    """DELETE rows from `table` where `column` = `value`."""
    import time
    for attempt in range(5):
        try:
            resp = requests.delete(
                f"{url}/rest/v1/{table}",
                headers=_headers(key),
                params={column: f"eq.{value}"},
                timeout=30,
            )
            if resp.ok:
                return
            err = f"{resp.status_code} {resp.text}"
        except Exception as exc:
            err = str(exc)
        time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"Delete failed on {table}: {err}")


def supabase_delete_all(url, key, table):
    """DELETE every row in `table` (used for a full hierarchy resync)."""
    # "id=gt.0" only works for integer primary keys and silently matches
    # nothing on a uuid PK. "not.is.null" is type-agnostic.
    resp = requests.delete(
        f"{url}/rest/v1/{table}",
        headers=_headers(key, prefer="return=minimal,count=exact"),
        params={"id": "not.is.null"},
    )
    if not resp.ok:
        raise RuntimeError(f"Delete-all failed on {table}: {resp.status_code} {resp.text}")


def supabase_insert(url, key, table, rows, upsert_on=None):
    """INSERT rows in chunks. If upsert_on is given, does an upsert instead
    (on-conflict merge) using that column/constraint name."""
    if not rows:
        return 0
    prefer = "return=minimal"
    params = {}
    if upsert_on:
        prefer = f"resolution=merge-duplicates,return=minimal"
        params["on_conflict"] = upsert_on

    import time
    total = 0
    for i in range(0, len(rows), CHUNK_SIZE):
        chunk = rows[i:i + CHUNK_SIZE]
        success = False
        last_error = None
        for attempt in range(5):
            try:
                resp = requests.post(
                    f"{url}/rest/v1/{table}",
                    headers=_headers(key, prefer=prefer),
                    params=params,
                    json=chunk,
                    timeout=30,
                )
                if resp.ok:
                    success = True
                    break
                else:
                    last_error = f"{resp.status_code} {resp.text}"
            except Exception as exc:
                last_error = str(exc)
            time.sleep(2 * (attempt + 1))

        if not success:
            raise RuntimeError(
                f"Insert failed on {table} (rows {i}-{i+len(chunk)}): {last_error}"
            )
        total += len(chunk)
    return total


PAGE_SIZE = 1000  # PostgREST's own default hard cap per request


def supabase_select(url, key, table, select="*", params=None):
    """SELECT rows, following pagination until the table is exhausted.

    IMPORTANT: PostgREST caps every response at 1000 rows. The previous
    version issued a single un-paginated request, so the de-duplication
    checks below only ever saw the first 1000 existing rows. Once a batch
    grew past that — which every real month does — re-running an upload
    re-inserted everything beyond row 1000 as duplicates. This is the main
    cause of doubled visits and coaching days on the dashboard.
    """
    p = {"select": select}
    if params:
        p.update(params)

    rows = []
    offset = 0
    while True:
        headers = _headers(key)
        headers["Range-Unit"] = "items"
        headers["Range"] = f"{offset}-{offset + PAGE_SIZE - 1}"
        resp = requests.get(f"{url}/rest/v1/{table}", headers=headers, params=p)
        if not resp.ok:
            raise RuntimeError(f"Select failed on {table}: {resp.status_code} {resp.text}")
        page = resp.json()
        rows.extend(page)
        if len(page) < PAGE_SIZE:
            return rows
        offset += PAGE_SIZE


# ── cleaning helpers ────────────────────────────────────────────────────────

def clean_code(value):
    """Normalise an employee code: '13512.0' / 13512.0 / '13512' → '13512'."""
    if value is None:
        return None
        
    if isinstance(value, float):
        if pd.isna(value):
            return None
        if value.is_integer():
            value = int(value)
            
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None
    if s.endswith(".0"):
        s = s[:-2]
    return s or None


def clean_str(value):
    if value is None:
        return None
        
    if isinstance(value, float):
        if pd.isna(value):
            return None
        if value.is_integer():
            value = int(value)
            
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None
    return s


def clean_num(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        v = float(value)
        if v.is_integer():
            return int(v)
        return v
    except (ValueError, TypeError):
        return None


def hm_to_minutes(hm_str):
    """Convert an 'h:mm' string (e.g. '1:23') back to total minutes for storage."""
    s = clean_str(hm_str)
    if not s or ":" not in s:
        return None
    try:
        h, m = s.split(":")
        return int(h) * 60 + int(m)
    except ValueError:
        return None


def hm_to_decimal_hours(hm_str):
    """Convert an 'h:mm' string (e.g. '1:23') to decimal hours for dashboard front-end."""
    s = clean_str(hm_str)
    if not s or ":" not in s:
        return None
    try:
        h, m = s.split(":")
        return int(h) + int(m) / 60.0
    except ValueError:
        return None


def detect_period(df):
    """Given a DataFrame with a 'date' column, return the dominant month
    as e.g. 'June 2026', matching pulpoplus_rebuild_summary's own logic."""
    month_counter = Counter()
    for d in df.get("date", []):
        if not d: continue
        d_clean = clean_date(d)
        if not d_clean: continue
        try:
            dt = datetime.strptime(d_clean[:10], "%Y-%m-%d")
            month_counter[(dt.year, dt.month)] += 1
        except ValueError:
            continue
    if not month_counter:
        return None
    (year, month), _ = month_counter.most_common(1)[0]
    return datetime(year, month, 1).strftime("%B %Y")


def dedupe_rows(rows, key_fields):
    """Collapse rows that share the same natural key, keeping the LAST
    occurrence (later rows are assumed to be corrections). This guards
    against duplicate entries within a single sheet AND against the same
    session appearing in more than one workbook uploaded in the same batch."""
    deduped = {}
    for row in rows:
        k = tuple(row.get(f) for f in key_fields)
        deduped[k] = row  # later occurrence overwrites earlier one
    return list(deduped.values())


# ── hierarchy sync ──────────────────────────────────────────────────────────

def sync_hierarchy(url, key, hierarchy_path, debug=False):
    print(f"\nLoading hierarchy from {hierarchy_path} ...")
    hmap = hier.load_hierarchy(hierarchy_path)
    teams = hmap.teams()
    print(f"   {len(hmap.all_employees)} people across {len(teams)} teams")

    # 1. Upsert teams (unique on name), then fetch back their generated IDs.
    team_rows = []
    for t in teams:
        blm = next((e for e in hmap.all_employees if e.team == t and e.role == "BLM"), None)
        team_rows.append({
            "name": t,
            "blm_name": blm.name if blm else None,
            "blm_code": blm.code if blm else None,
        })
    print(f"\nUpserting {len(team_rows)} team(s)...")
    supabase_insert(url, key, "teams", team_rows, upsert_on="name")

    existing_teams = supabase_select(url, key, "teams", select="id,name")
    team_id_by_name = {t["name"]: t["id"] for t in existing_teams}

    # 2. Upsert `hierarchy` instead of full replace to prevent breaking foreign keys
    #    from the `app_users` table.
    print("\nFetching existing hierarchy rows...")
    existing_hierarchy = supabase_select(url, key, "hierarchy", select="id, employee_name, team_id")
    
    # Map (employee_name, team_id) to list of available IDs
    ids_by_key = {}
    for r in existing_hierarchy:
        name = r.get("employee_name")
        tid = r.get("team_id")
        if name:
            k = (name, tid)
            if k not in ids_by_key:
                ids_by_key[k] = []
            ids_by_key[k].append(r["id"])
            # Fallback by name alone if team_id didn't match
            name_k = (name, None)
            if name_k not in ids_by_key:
                ids_by_key[name_k] = []
            ids_by_key[name_k].append(r["id"])

    insert_rows = []
    update_rows = []
    seen_keys = set()
    used_ids = set()

    for i, e in enumerate(hmap.all_employees):
        tid = team_id_by_name.get(e.team)
        key_tuple = (e.name, tid)
        if not e.name or key_tuple in seen_keys:
            continue
        seen_keys.add(key_tuple)
        
        row = {
            "team_id": tid,
            "employee_name": e.name,
            "employee_code": e.code or None,
            "role": e.role,
            "division_name": e.territory or None,
            "supervisor_name": e.supervisor or e.area_manager or None,
            "area_manager_name": e.area_manager or None,
            "row_number": i,
        }
        
        # Match existing ID by (name, team_id) first, then by (name, None)
        matched_id = None
        for candidate_key in [key_tuple, (e.name, None)]:
            if candidate_key in ids_by_key:
                for cand_id in ids_by_key[candidate_key]:
                    if cand_id not in used_ids:
                        matched_id = cand_id
                        used_ids.add(cand_id)
                        break
            if matched_id:
                break
                
        if matched_id:
            row["id"] = matched_id
            update_rows.append(row)
        else:
            insert_rows.append(row)

    print(f"Upserting {len(update_rows)} existing and inserting {len(insert_rows)} new hierarchy row(s)...")
    inserted = 0
    if update_rows:
        inserted += supabase_insert(url, key, "hierarchy", update_rows, upsert_on="id")
    if insert_rows:
        inserted += supabase_insert(url, key, "hierarchy", insert_rows)
    print(f"   OK Hierarchy synced: {inserted} rows")


# ── workbook upload ──────────────────────────────────────────────────────────

def upload_workbook(url, key, workbook_path, period=None, batch=None, append=False, debug=False):
    print(f"\nReading {workbook_path} ...")
    sheets = pd.read_excel(workbook_path, sheet_name=None)

    raw_df = sheets.get("Raw Data")
    if raw_df is None:
        print("No 'Raw Data' sheet found in this workbook.")
        sys.exit(1)

    if not period:
        # Standardise columns just like the extractor does so we can read Raw CRM exports directly
        # as well as already-rebuilt summaries.
        col_map = {
            "Employee Code": "user_code",
            "Employee Name": "user",
            "Date": "date",
            "Time": "time",
            "Shift": "shift",
            "Account Code": "acc_id",
            "Account Name": "acc_name",
            "Doctor Code": "doctor_key",
            "Doctor Name": "doctor_name",
            "Specialty": "specialty",
            "Classification": "classification",
            "Products": "products",
            "Visit Type": "visit_type_raw",
            "Account Type": "acc_type_raw",
            "Territory": "territory",
            "Team": "team",
            "Notes": "notes",
            "User": "user",
        }
        raw_df = raw_df.rename(columns=col_map)
        # Also ensure we catch lowercase variants if they were somehow missed
        raw_df.columns = [str(c).lower().strip() for c in raw_df.columns]

        period = detect_period(raw_df)
        print(f"   Auto-detected period: {period}")
    batch = batch or period or "unbatched"

    # ---- Raw Data → visits ----
    if not append:
        print(f"\nClearing existing 'visits' rows for batch '{batch}'...")
        supabase_delete_where(url, key, "visits", "upload_batch", batch)

    visit_rows = []
    for _, r in raw_df.iterrows():
        visit_rows.append({
            "team":                 clean_str(r.get("team")),
            "user":                 clean_str(r.get("user")),
            "employee_code":        clean_code(r.get("user_code") or r.get("employee code")),
            "territory":            clean_str(r.get("territory")),
            "visit_date":           clean_date(r.get("date")),
            "visit_time":           clean_time(r.get("time")),
            "acc_type_raw":         clean_str(r.get("acc_type_raw")),
            "acc_type_category":    clean_str(r.get("acc_type_category")),
            "shift":                clean_str(r.get("shift")),
            "visit_type_raw":       clean_str(r.get("visit_type_raw")),
            "visit_type_category":  clean_str(r.get("visit_type_category")),
            "acc_id":               clean_str(r.get("acc_id") or r.get("acc_name")),
            "acc_name":             clean_str(r.get("acc_name")),
            "doctor_key":           clean_str(r.get("doctor_key") or r.get("doctor_name")),
            "doctor_name":          clean_str(r.get("doctor_name")),
            "specialty":            clean_str(r.get("specialty")),
            "classification":       clean_str(r.get("classification")),
            "products":             clean_str(r.get("products")),
            "notes":                clean_str(r.get("notes")),
            "upload_batch":         batch,
            "period":               period,   # e.g. "July 2026" — enables fast period-based queries
        })

    if append:
        print("Fetching existing visits to prevent duplicates...")
        existing = supabase_select(url, key, "visits", select="user,visit_date,visit_time,acc_id", params={"upload_batch": f"eq.{batch}"})
        existing_keys = {(r.get("user"), r.get("visit_date"), r.get("visit_time"), r.get("acc_id")) for r in existing}
        before = len(visit_rows)
        visit_rows = [r for r in visit_rows if (r["user"], r["visit_date"], r["visit_time"], r["acc_id"]) not in existing_keys]
        if len(visit_rows) < before:
            print(f"   Skipped {before - len(visit_rows)} existing visit row(s)")

    print(f"Inserting {len(visit_rows)} visit row(s)...")
    n = supabase_insert(url, key, "visits", visit_rows)
    print(f"   OK visits: {n} rows")

    # ---- Coaching Days ----
    coaching_df = sheets.get("Coaching Days")
    if coaching_df is not None and not coaching_df.empty:
        if not append:
            print(f"\nClearing existing 'coaching_days' rows for batch '{batch}'...")
            supabase_delete_where(url, key, "coaching_days", "upload_batch", batch)

        coaching_rows = []
        for _, r in coaching_df.iterrows():
            coaching_rows.append({
                "manager_name":    clean_str(r.get("Manager")),
                "manager_code":    clean_code(r.get("Manager Code")),
                "rep_name":        clean_str(r.get("Rep")),
                "rep_code":        clean_code(r.get("Rep Code")),
                "coaching_date":   clean_date(r.get("Date")),
                "team":            clean_str(r.get("Team")),
                "am_visits":       clean_num(r.get("AM Visits")),
                "am_accompanied":  clean_num(r.get("AM Accompanied")),
                "pm_visits":       clean_num(r.get("PM Visits")),
                "pm_accompanied":  clean_num(r.get("PM Accompanied")),
                "upload_batch":    batch,
                "period":          period,
            })

        # De-duplicate on the natural key of a coaching session, so the
        # same manager/rep/date pair can never end up as two rows — whether
        # the duplicate came from the same sheet or from a second workbook
        # uploaded into the same batch.
        before = len(coaching_rows)
        coaching_rows = dedupe_rows(
            coaching_rows,
            key_fields=("manager_name", "rep_name", "coaching_date", "team"),
        )
        if len(coaching_rows) < before:
            print(f"   Removed {before - len(coaching_rows)} duplicate coaching row(s) internally")

        if append:
            print("Fetching existing coaching days to prevent duplicates...")
            existing = supabase_select(url, key, "coaching_days", select="manager_name,rep_name,coaching_date", params={"upload_batch": f"eq.{batch}"})
            existing_keys = {(r.get("manager_name"), r.get("rep_name"), r.get("coaching_date")) for r in existing}
            before = len(coaching_rows)
            coaching_rows = [r for r in coaching_rows if (r["manager_name"], r["rep_name"], r["coaching_date"]) not in existing_keys]
            if len(coaching_rows) < before:
                print(f"   Skipped {before - len(coaching_rows)} existing coaching day row(s)")

        print(f"Inserting {len(coaching_rows)} coaching day row(s)...")
        n = supabase_insert(url, key, "coaching_days", coaching_rows)
        print(f"   OK coaching_days: {n} rows")
    else:
        print("\n   (no Coaching Days sheet/rows found — skipping)")

    # ---- Summary → summaries ----
    summary_df = sheets.get("Summary")
    if summary_df is not None and not summary_df.empty:
        if not append:
            print(f"\nClearing existing 'summaries' rows for batch '{batch}'...")
            supabase_delete_where(url, key, "summaries", "upload_batch", batch)

        summary_rows = []
        for _, r in summary_df.iterrows():
            summary_rows.append({
                "employee_code":        clean_code(r.get("Employee Code")),
                "user_name":            clean_str(r.get("User")),
                "period":               period,
                "team":                 clean_str(r.get("Team")),
                "territory":            clean_str(r.get("Territory")),
                "is_manager":           str(r.get("Is Manager")).strip().lower() == "yes",
                "working_days":         clean_num(r.get("Working Days")),
                "complete_field_days":  clean_num(r.get("Complete Field Days")),
                "office_work_days":     clean_num(r.get("Office Work Days")),
                "no_activities":        clean_num(r.get("No. of Activities")),
                "no_events":            clean_num(r.get("No. of Events")),
                "double_visit_days":    clean_num(r.get("Double Visit Days")),
                "coaching_days":        clean_num(r.get("Coaching Days")),
                "am_shift_days":        clean_num(r.get("AM Shift Days")),
                "pm_shift_days":        clean_num(r.get("PM Shift Days")),
                "avg_am_duration_min":  hm_to_minutes(r.get("Avg AM Shift Duration (h:mm)")),
                "avg_pm_duration_min":  hm_to_minutes(r.get("Avg PM Shift Duration (h:mm)")),
                "am_calls":             clean_num(r.get("AM Calls")),
                "pm_calls":             clean_num(r.get("PM Calls")),
                "am_call_rate":         clean_num(r.get("AM Call Rate")),
                "pm_call_rate":         clean_num(r.get("PM Call Rate")),
                "avg_am_start_time":    clean_str(r.get("Avg AM Starting Time")),
                "am_unique_doctors":    clean_num(r.get("AM Unique Doctors")),
                "pm_unique_doctors":    clean_num(r.get("PM Unique Doctors")),
                "total_am_covered":     clean_num(r.get("Total AM Covered")),
                "total_pm_covered":     clean_num(r.get("Total PM Covered")),
                "clinic_covered":       clean_num(r.get("Clinic Doctors Covered")),
                "polyclinic_covered":   clean_num(r.get("Poly Clinic Doctors Covered")),
                "amcenter_covered":     clean_num(r.get("AM Center Doctors Covered")),
                "hospital_covered":     clean_num(r.get("Hospital Doctors Covered")),
                "pharmacies_visited":   clean_num(r.get("Pharmacies Visited")),
                "pharmacies_covered":   clean_num(r.get("Pharmacies Covered")),
                "total_product_calls":  clean_num(r.get("Total Product Calls")),
                "distinct_products":    clean_num(r.get("Distinct Products Promoted")),
                "product_calls_detail": clean_str(r.get("Product Detail")),
                "avg_am_shift_hm":      hm_to_decimal_hours(r.get("Avg AM Shift Duration (h:mm)")),
                "avg_pm_shift_hm":      hm_to_decimal_hours(r.get("Avg PM Shift Duration (h:mm)")),
                "avg_field_overall_hm":  hm_to_decimal_hours(r.get("Avg Field Duration Overall (h:mm)")),
                "total_visits":         clean_num(r.get("Total Visits")),
                "upload_batch":         batch,
            })

        if append:
            print("Fetching existing summaries to prevent duplicates...")
            existing = supabase_select(url, key, "summaries", select="user_name,employee_code", params={"upload_batch": f"eq.{batch}"})
            existing_keys = {(r.get("user_name"), r.get("employee_code")) for r in existing}
            before = len(summary_rows)
            summary_rows = [r for r in summary_rows
                            if (r["user_name"], r["employee_code"]) not in existing_keys]
            if len(summary_rows) < before:
                print(f"   Skipped {before - len(summary_rows)} existing summary row(s)")

        print(f"Inserting {len(summary_rows)} summary row(s)...")
        n = supabase_insert(url, key, "summaries", summary_rows)
        print(f"   OK summaries: {n} rows")
    else:
        print("\n   (no Summary sheet/rows found — skipping)")

    # ---- Specialty x Class ----
    spec_df = sheets.get("Specialty x Class")
    if spec_df is not None and not spec_df.empty:
        if not append:
            print(f"\nClearing existing 'specialty_classification' rows for batch '{batch}'...")
            supabase_delete_where(url, key, "specialty_classification", "upload_batch", batch)

        spec_rows = []
        for _, r in spec_df.iterrows():
            spec_rows.append({
                "employee_code":  clean_code(r.get("Employee Code")),
                "user_name":      clean_str(r.get("User")),
                "period":         period,
                "specialty":      clean_str(r.get("Specialty")),
                "classification": clean_str(r.get("Classification")),
                "shift":          clean_str(r.get("Shift")),
                "call_count":     clean_num(r.get("Call Count")),
                "unique_doctors": clean_num(r.get("Unique Doctors")),
                "upload_batch":   batch,
            })

        if append:
            print("Fetching existing specialty_classification to prevent duplicates...")
            existing = supabase_select(url, key, "specialty_classification", select="user_name,specialty,classification,shift", params={"upload_batch": f"eq.{batch}"})
            existing_keys = {(r.get("user_name"), r.get("specialty"), r.get("classification"), r.get("shift")) for r in existing}
            before = len(spec_rows)
            spec_rows = [r for r in spec_rows if (r["user_name"], r["specialty"], r["classification"], r["shift"]) not in existing_keys]
            if len(spec_rows) < before:
                print(f"   Skipped {before - len(spec_rows)} existing specialty_classification row(s)")

        print(f"Inserting {len(spec_rows)} specialty-classification row(s)...")
        n = supabase_insert(url, key, "specialty_classification", spec_rows)
        print(f"   OK specialty_classification: {n} rows")
    else:
        print("\n   (no Specialty x Class sheet/rows found — skipping)")

    # ---- Product Calls per spec ----
    prod_df = sheets.get("Product Calls per spec")
    if prod_df is not None and not prod_df.empty:
        if not append:
            print(f"\nClearing existing 'product_calls' rows for batch '{batch}'...")
            supabase_delete_where(url, key, "product_calls", "upload_batch", batch)

        prod_rows = []
        for _, r in prod_df.iterrows():
            prod_rows.append({
                "employee_code": clean_code(r.get("Employee Code")),
                "user_name":     clean_str(r.get("User")),
                "period":        period,
                "specialty":     clean_str(r.get("Specialty")),
                "product":       clean_str(r.get("Product")),
                "shift":         clean_str(r.get("Shift")),
                "call_count":    clean_num(r.get("Call Count")),
                "unique_doctors": clean_num(r.get("Unique Doctors")),
                "upload_batch":  batch,
            })

        if append:
            print("Fetching existing product_calls to prevent duplicates...")
            existing = supabase_select(url, key, "product_calls", select="user_name,specialty,product,shift", params={"upload_batch": f"eq.{batch}"})
            existing_keys = {(r.get("user_name"), r.get("specialty"), r.get("product"), r.get("shift")) for r in existing}
            before = len(prod_rows)
            prod_rows = [r for r in prod_rows if (r["user_name"], r["specialty"], r["product"], r["shift"]) not in existing_keys]
            if len(prod_rows) < before:
                print(f"   Skipped {before - len(prod_rows)} existing product_calls row(s)")

        print(f"Inserting {len(prod_rows)} product-call row(s)...")
        n = supabase_insert(url, key, "product_calls", prod_rows)
        print(f"   OK product_calls: {n} rows")
    else:
        print("\n   (no Product Calls per spec sheet/rows found — skipping)")

    print(f"\nDone — batch '{batch}' uploaded.")


# ── entry point ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Upload a rebuilt PulpoPlus workbook and/or hierarchy file to Supabase."
    )
    parser.add_argument("--workbook", default=None, help="Rebuilt summary .xlsx to upload")
    parser.add_argument("--hierarchy", default=None, help="Team structure .xlsx to sync")
    parser.add_argument("--period", default=None, help="Period label, e.g. 'June 2026' (auto-detected if omitted)")
    parser.add_argument("--batch", default=None, help="Upload batch id (defaults to the period)")
    parser.add_argument("--append", action="store_true", help="Append data without deleting existing rows in the batch")
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()

    if not args.workbook and not args.hierarchy:
        parser.error("Provide --workbook and/or --hierarchy (at least one is required)")

    url, key = _supabase_config()

    print("\n" + "=" * 65)
    print("PulpoPlus → Supabase Uploader")
    print("=" * 65)

    if args.hierarchy:
        sync_hierarchy(url, key, args.hierarchy, debug=args.debug)

    if args.workbook:
        upload_workbook(url, key, args.workbook, period=args.period,
                        batch=args.batch, append=args.append, debug=args.debug)


if __name__ == "__main__":
    main()