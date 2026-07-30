#!/usr/bin/env python3
"""
pulpoplus_extract_visits.py
────────────────────────────
Reads a raw "Actual Visits Report" export from PulpoPlus (the messy HTML
file that's saved with a .xls extension — the same kind of file the
extractor itself scrapes) and builds a full summary workbook (Summary,
Specialty x Class, Product Calls per spec, Coaching Days, Raw Data) using
the exact same calculation logic as pulpoplus_rebuild_summary.py.

Handles both kinds of raw report:
  • Team-wide export (caption has no name) — one row can list several
    people in "Members"; each gets their own credited record.
  • Single-person "my visits" export (caption names one person, e.g.
    "Adel Morsy") — every row belongs to that person only; the row is
    NOT exploded per member even if others (reps) are listed with them,
    since only one visit actually happened.

Sections parsed: Visits, Pharmacies Visits, Office Work, Activities, Events.
Each section's own "#" column is the row-sequence counter the PulpoPlus UI
itself uses, and is a good sanity check against the totals in this file.

Usage:
    python pulpoplus_extract_visits.py visits.xls
    python pulpoplus_extract_visits.py visits.xls --out summary.xlsx --team "PLATINUM 1"
    python pulpoplus_extract_visits.py visits.xls --debug
"""

import os
import re
import sys
import html
import argparse
from collections import OrderedDict

# Reuse all the calculation logic from the rebuild script unchanged.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pulpoplus_rebuild_summary as reb


# ── low-level HTML row/cell parsing ──────────────────────────────────────────

def _cellsplit(row_html):
    """Split one <tr>...row html (without the leading '<tr') into its <td> cells."""
    parts = re.split(r"<td[^>]*>", row_html)
    return parts[1:]  # part[0] is the <tr ...> attribute soup, not a cell


def _clean_cell(c):
    c = re.sub(r"<br\s*/?>", "\n", c)
    c = re.sub(r"<[^>]+>", "", c)
    c = html.unescape(c)
    return c.strip()


def _parse_dash_list(raw):
    """'-Foo\n-Bar   \n-Bar' -> ['Foo', 'Bar'] (order-preserving de-dupe)."""
    if not raw:
        return []
    out = OrderedDict()
    for line in raw.split("\n"):
        line = line.strip()
        if line.startswith("-"):
            line = line[1:]
        line = line.strip()
        if line:
            out[line] = True
    return list(out.keys())


def _parse_products(raw):
    items = _parse_dash_list(raw)
    return ", ".join(items)


def _split_date_time(raw):
    lines = [l.strip() for l in raw.split("\n") if l.strip()]
    date = lines[0] if len(lines) >= 1 else ""
    time = lines[1] if len(lines) >= 2 else ""
    return date, time


# ── report-owner detection ───────────────────────────────────────────────────

def detect_report_owner(data):
    """
    The caption looks like either:
      '<br/>Actual  Visits Report <br/>  <br/>From 2026-06-01 To 2026-06-15'
        -> no name -> team-wide report -> returns None
      '<br/>Actual  Visits Report <br/> Adel Morsy <br/>From 2026-07-01 To 2026-07-14'
        -> returns 'Adel Morsy'
    """
    m = re.search(r"<caption.*?>(.*?)</caption>", data, re.S)
    if not m:
        return None
    caption_html = m.group(1)
    parts = re.split(r"<br\s*/?>", caption_html)
    parts = [_clean_cell(p) for p in parts]
    parts = [p for p in parts if p]
    for p in parts:
        if p.lower().startswith("actual"):
            continue
        if p.lower().startswith("from") and " to " in p.lower():
            continue
        return p  # first remaining non-empty line is the report owner's name
    return None


# ── section extraction ───────────────────────────────────────────────────────

def _all_sections(data):
    """Return OrderedDict of {legend_name: section_html} in document order."""
    matches = list(re.finditer(r"<legend>(.*?)</legend>", data, re.S))
    sections = OrderedDict()
    for i, m in enumerate(matches):
        name = _clean_cell(m.group(1))
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(data)
        sections[name] = data[start:end]
    return sections


_HEADER_WORDS = ("territory", "acc. type", "acc type", "doctor", "specialty",
                 "classification", "products", "members", "date", "visit type",
                 "notes", "feedback", "pharmacy", "activity", "event")


def _looks_like_header(row_html):
    """True if this <tr> is a header rather than a data row."""
    if "<th" in row_html.lower():
        return True
    cells = [_clean_cell(c).lower() for c in _cellsplit(row_html)]
    filled = [c for c in cells if c]
    if not filled:
        return False
    hits = sum(1 for c in filled if any(w in c for w in _HEADER_WORDS))
    return hits >= max(2, len(filled) // 3)


def _rows_of(section_html):
    """Data rows of a section, with any leading header row removed.

    The old version did an unconditional [1:] on every row containing a
    <td>. Header cells are usually <th>, so they were already excluded by
    that filter — meaning the [1:] silently discarded the FIRST REAL DATA
    ROW of every section. Now the first row is only dropped if it actually
    looks like a header.
    """
    rows = [r for r in section_html.split("<tr") if "<td" in r]
    if rows and _looks_like_header(rows[0]):
        return rows[1:]
    return rows


# ── per-section row builders ──────────────────────────────────────────────────

def _emit(records, base, members, owner):
    """Append one or more records for a parsed row, respecting single-owner mode."""
    if owner:
        rec = dict(base)
        rec["user"] = owner
        rec["members"] = members if members else [owner]
        records.append(rec)
    else:
        if not members:
            return
        for member in members:
            rec = dict(base)
            rec["user"] = member
            rec["members"] = list(members)
            records.append(rec)


def _parse_visits_section(rows, team, owner, records, debug, label="Visits"):
    skipped = 0
    for row in rows:
        cells = [_clean_cell(c) for c in _cellsplit(row)]
        if len(cells) != 17:
            skipped += 1
            continue

        (_blank, _seq, territory_raw, visit_type_raw, acc_type_raw, acc_id,
         acc_name, doctor_key, doctor_name, specialty, classification,
         date_time_raw, products_raw, _feedback_raw, members_raw,
         notes_raw, _visit_root) = cells

        territory = territory_raw.split("\n")[0].strip()
        date, time = _split_date_time(date_time_raw)
        products = _parse_products(products_raw)
        members = _parse_dash_list(members_raw)

        acc_type_category = reb.ACC_TYPE_MAP.get(acc_type_raw.upper(), acc_type_raw)
        shift = reb.SHIFT_MAP.get(acc_type_category, "")
        visit_type_category = {"S": "Single", "D": "Double"}.get(
            visit_type_raw.upper(), visit_type_raw
        )

        base = {
            "team": team, "territory": territory, "date": date, "time": time,
            "acc_type_raw": acc_type_raw, "acc_type_category": acc_type_category,
            "shift": shift, "visit_type_raw": visit_type_raw,
            "visit_type_category": visit_type_category, "acc_id": acc_id,
            "acc_name": acc_name, "doctor_key": doctor_key, "doctor_name": doctor_name,
            "notes": notes_raw, "specialty": specialty, "classification": classification,
            "products": products,
        }
        _emit(records, base, members, owner)

    # Skipped rows used to be reported only under --debug, so a column-count
    # change in the export dropped records with no visible sign.
    if skipped:
        print(f"  ! {label}: {skipped} of {len(rows)} row(s) skipped — "
              f"unexpected column count (export format may have changed)")
    elif debug:
        print(f"  {label}: {len(rows)} rows parsed, 0 skipped")
    return len(rows)


def _parse_pharmacy_section(rows, team, owner, records, debug):
    skipped = 0
    for row in rows:
        cells = [_clean_cell(c) for c in _cellsplit(row)]
        if len(cells) != 18:
            skipped += 1
            continue

        (_blank, _seq, territory_raw, visit_type_raw, acc_type_raw, acc_id,
         acc_name, doctor_key, doctor_name, specialty, classification,
         date_time_raw, products_raw, _last_order, _current_stock,
         _current_order, members_raw, notes_raw) = cells

        territory = territory_raw.split("\n")[0].strip()
        date, time = _split_date_time(date_time_raw)
        products = _parse_products(products_raw)
        members = _parse_dash_list(members_raw)

        acc_type_category = reb.ACC_TYPE_MAP.get(acc_type_raw.upper(), acc_type_raw)
        shift = reb.SHIFT_MAP.get(acc_type_category, "")
        visit_type_category = {"S": "Single", "D": "Double"}.get(
            visit_type_raw.upper(), visit_type_raw
        )

        base = {
            "team": team, "territory": territory, "date": date, "time": time,
            "acc_type_raw": acc_type_raw, "acc_type_category": acc_type_category,
            "shift": shift, "visit_type_raw": visit_type_raw,
            "visit_type_category": visit_type_category, "acc_id": acc_id,
            "acc_name": acc_name, "doctor_key": doctor_key, "doctor_name": doctor_name,
            "notes": notes_raw, "specialty": specialty, "classification": classification,
            "products": products,
        }
        _emit(records, base, members, owner)

    if skipped:
        print(f"  ! Pharmacies Visits: {skipped} of {len(rows)} row(s) skipped — "
              f"unexpected column count (export format may have changed)")
    elif debug:
        print(f"  Pharmacies Visits: {len(rows)} rows parsed, 0 skipped")
    return len(rows)


def _parse_ow_activity_section(rows, team, owner, records, debug, acc_type_raw, acc_type_category, label):
    skipped = 0
    for row in rows:
        cells = [_clean_cell(c) for c in _cellsplit(row)]
        if len(cells) != 9:
            skipped += 1
            continue

        (_seq, _id, territory_raw, _type_raw, shift, date, time,
         members_raw, comments_raw) = cells

        members = _parse_dash_list(members_raw)

        base = {
            "team": team, "territory": territory_raw.strip(), "date": date.strip(),
            "time": time.strip(), "acc_type_raw": acc_type_raw,
            "acc_type_category": acc_type_category, "shift": shift.strip(),
            "visit_type_raw": "", "visit_type_category": "", "acc_id": "",
            "acc_name": "", "doctor_key": "", "doctor_name": "", "notes": comments_raw,
            "specialty": "", "classification": "", "products": "",
        }
        _emit(records, base, members, owner)

    # Skipped rows used to be reported only under --debug, so a column-count
    # change in the export dropped records with no visible sign.
    if skipped:
        print(f"  ! {label}: {skipped} of {len(rows)} row(s) skipped — "
              f"unexpected column count (export format may have changed)")
    elif debug:
        print(f"  {label}: {len(rows)} rows parsed, 0 skipped")
    return len(rows)


def _parse_events_section(rows, team, owner, records, debug):
    skipped = 0
    for row in rows:
        cells = [_clean_cell(c) for c in _cellsplit(row)]
        if len(cells) < 9:
            skipped += 1
            continue

        # #, ID, Line, Type, Event, Date From, Date To, No. of Doctors, Nearest Pharmacy, Employee
        cells = cells + [""] * (10 - len(cells))
        (_seq, _id, line, ev_type, event, date_from, _date_to,
         _n_doctors, _nearest_pharmacy, employee) = cells[:10]

        member = employee.strip() or owner or ""
        base = {
            "team": team, "territory": line.strip(), "date": date_from.strip(),
            "time": "", "acc_type_raw": "E", "acc_type_category": "Events",
            "shift": "", "visit_type_raw": "", "visit_type_category": "",
            "acc_id": "", "acc_name": "", "doctor_key": "", "doctor_name": "",
            "notes": f"{ev_type} - {event}".strip(" -"), "specialty": "",
            "classification": "", "products": "",
        }
        _emit(records, base, [member] if member else [], owner or member)

    if skipped:
        print(f"  ! Events: {skipped} of {len(rows)} row(s) skipped — "
              f"unexpected column count (export format may have changed)")
    elif debug:
        print(f"  Events: {len(rows)} rows parsed, 0 skipped")
    return len(rows)


# ── de-duplication ───────────────────────────────────────────────────────────

def _dedupe_exact_records(records, debug=False):
    """
    PulpoPlus's raw export duplicates the row for 'Double' visits (a manager
    doing a joint/coaching call with a rep comes out as two <tr> rows for
    the SAME credited user, not one row per participant). Left alone, that
    double-credits everything downstream (Total Visits, PM/AM Calls,
    Product Calls, ...) for whoever the duplicated rows belong to.

    A record is a duplicate if it shares the same user, date, time, shift,
    and account/doctor — a real second visit by the same rep can never
    land on the exact same account at the exact same logged time on the
    same day. We deliberately don't require every field (acc_id, notes) to
    match too, since duplicated rows can carry a slightly different
    internal id or formatting despite describing the same real visit.
    """
    seen = set()
    deduped = []
    dropped = 0
    for rec in records:
        key = (
            rec["user"], rec["date"], rec["time"], rec["shift"],
            rec["acc_id"] or rec["doctor_key"] or rec["acc_name"],
        )
        if key in seen:
            dropped += 1
            continue
        seen.add(key)
        deduped.append(rec)

    if debug:
        print(f"  Deduped {dropped} duplicate record(s) "
              f"(commonly caused by 'Double' visit rows being exported twice)")
    return deduped


# ── build records ────────────────────────────────────────────────────────────

def parse_visits_report(path, team="", debug=False):
    with open(path, encoding="utf-8", errors="ignore") as f:
        data = f.read()

    owner = detect_report_owner(data)
    if debug:
        print(f"  Report owner detected: {owner!r} "
              f"({'single-user mode' if owner else 'team-wide mode (Members column drives user)'})")

    sections = _all_sections(data)
    records = []
    counts = {}

    for name, section_html in sections.items():
        rows = _rows_of(section_html)
        key = name.lower()
        if key == "visits":
            counts[name] = _parse_visits_section(rows, team, owner, records, debug)
        elif key == "pharmacies visits":
            counts[name] = _parse_pharmacy_section(rows, team, owner, records, debug)
        elif key == "office work":
            counts[name] = _parse_ow_activity_section(
                rows, team, owner, records, debug, "OW", "Office Work", "Office Work")
        elif key == "activities":
            counts[name] = _parse_ow_activity_section(
                rows, team, owner, records, debug, "A", "Activities", "Activities")
        elif key == "events":
            counts[name] = _parse_events_section(rows, team, owner, records, debug)
        else:
            if debug:
                print(f"  (skipping unrecognized section '{name}' — {len(rows)} rows)")

    filtered_records = []
    for r in records:
        if "-06-" in r.get("date", "") or "/06/" in r.get("date", ""):
            continue
        filtered_records.append(r)
    records = filtered_records

    pre_dedupe_count = len(records)
    records = _dedupe_exact_records(records, debug=debug)
    dropped = pre_dedupe_count - len(records)

    summary_bits = ", ".join(f"{n} {c}" for n, c in counts.items())
    print(f"  ✓ Built {pre_dedupe_count} raw-data records ({summary_bits})")
    if dropped:
        print(f"  ✓ Removed {dropped} exact-duplicate record(s) "
              f"from 'Double' visit rows exported twice")
    return records


def infer_team_from_filename(filename):
    base = os.path.basename(filename)
    stem = os.path.splitext(base)[0]
    # Remove date patterns like "July 2026 - " or "2026-07 - "
    clean_stem = re.sub(r'^(?:[A-Za-z]+\s+\d{4}|\d{4}[-_]\d{2})\s*[-_]?\s*', '', stem)
    # Remove suffix keywords like _rebuilt, _report, visits
    clean_stem = re.sub(r'[-_](?:rebuilt|report|visits|export)$', '', clean_stem, flags=re.I)
    clean_stem = clean_stem.strip()
    if clean_stem and clean_stem.lower() not in {"visits", "report", "actual visits report"}:
        return clean_stem.title()
    return ""


# ── entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Parse a raw PulpoPlus 'Actual Visits Report' (.xls HTML export) "
                    "and build a full summary workbook."
    )
    parser.add_argument("input", help="Path to the raw visits report (.xls)")
    parser.add_argument("--out", default=None, help="Output .xlsx path")
    parser.add_argument("--team", default="", help="Team name to stamp on every record (optional)")
    parser.add_argument("--hierarchy", default=None,
                        help="Path to the 2026 Team Structure .xlsx "
                             "(auto-discovered if omitted; skipped if --team is given)")
    parser.add_argument("--no-hierarchy", action="store_true",
                        help="Skip hierarchy file even if one is found")
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()

    out_path = args.out  # resolved after records load if not explicitly given

    print("\n" + "=" * 65)
    print("PulpoPlus Visits Report Extractor")
    print("=" * 65)
    print(f"\n📂 Input:  {args.input}")
    if out_path:
        print(f"📄 Output: {out_path}\n")
    else:
        print(f"📄 Output: (auto-named from data — determined after loading)\n")

    # ── hierarchy ─────────────────────────────────────────────────────────────
    hmap = None
    if not args.no_hierarchy:
        print("🏢 Loading 2026 team hierarchy...")
        hmap = reb._load_hierarchy_if_available(args.hierarchy)
        if hmap is None:
            print("   (no hierarchy file found — use --team or place "
                  "team_structure_2026.xlsx next to this script)")

    print("📥 Parsing raw HTML visits report...")
    records = parse_visits_report(args.input, team=args.team, debug=args.debug)

    # Enrich team from hierarchy (only for records still missing a team)
    if hmap is not None:
        print("\n🔗 Enriching records with hierarchy team assignments...")
        enriched = hmap.enrich_records(records, warn=args.debug)
        print(f"   ✓ {enriched} record(s) assigned a team from hierarchy")

    # Auto-infer team from filename if records still lack team assignments
    unassigned = sum(1 for r in records if not r.get("team"))
    if unassigned > 0:
        inferred = infer_team_from_filename(args.input)
        if inferred:
            for r in records:
                if not r.get("team"):
                    r["team"] = inferred
            print(f"   ✓ Inferred Team = '{inferred}' from filename for {unassigned} record(s)")

    if not out_path:
        filename = reb.auto_output_filename(records, fallback_base=os.path.splitext(os.path.basename(args.input))[0])
        recent_dir = os.environ.get("PERIOD_RECENT_DIR", r"E:\crm extractor\Periods\recent")
        if os.path.isdir(recent_dir):
            out_path = os.path.join(recent_dir, filename)
        else:
            out_path = filename
        print(f"   📄 Auto-named output destination: {out_path}")

    print("\n🌍 Backfilling territory...")
    reb.backfill_territory(records, debug=args.debug)

    print("\n👔 Identifying managers...")
    managers = reb.identify_managers(records)
    print(f"   Managers found ({len(managers)}): {sorted(managers)}")

    print("\n🤝 Computing coaching days...")
    coaching_days_by_manager, coaching_detail_df = reb.compute_coaching_days(
        records, managers, debug=args.debug
    )
    for mgr, days in sorted(coaching_days_by_manager.items()):
        print(f"   {mgr}: {days} coaching day(s)")

    print("\n📊 Computing per-user-per-day summary...")
    summary_df = reb.compute_summary(
        records,
        managers=managers,
        coaching_days_by_manager=coaching_days_by_manager,
        coaching_detail_df=coaching_detail_df,
        debug=args.debug,
        group_by_date=True,
    )

    print("\n🕐 Computing timing report (last visit per day)...")
    timing_df = reb.compute_timing_report(records)

    print("\n🔬 Computing specialty × classification...")
    spec_df = reb.compute_specialty_classification(records)

    print("\n💊 Computing product calls per specialty...")
    prod_df = reb.compute_product_calls(records)

    reb.save_output(
        summary_df, records, out_path,
        specialty_class_df=spec_df,
        product_df=prod_df,
        coaching_detail_df=coaching_detail_df,
        hmap=hmap,
        timing_df=timing_df,
    )


if __name__ == "__main__":
    main()