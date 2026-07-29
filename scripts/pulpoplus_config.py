"""
pulpoplus_config.py — shared credential loading and data-cleaning helpers.

Every script that talks to Supabase imports from here, so credentials live in
exactly one place (a .env file, never the source) and date handling is
consistent across the whole pipeline.
"""

import os
import sys
import base64
import json
from pathlib import Path
from datetime import datetime, date

# ── .env loading (no third-party dependency required) ──────────────────────


def _load_dotenv():
    """Read a .env file from this folder or its parent into os.environ.

    Values already set in the real environment always win, so a PowerShell
    `$env:` override still works.
    """
    here = Path(__file__).resolve().parent
    for candidate in (here / ".env", here.parent / ".env"):
        if not candidate.is_file():
            continue
        for line in candidate.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            k, v = k.strip(), v.strip().strip('"').strip("'")
            os.environ.setdefault(k, v)
        break


_load_dotenv()


def _required(name, hint=""):
    val = (os.environ.get(name) or "").strip()
    if not val:
        print(f"\nMissing required environment variable: {name}")
        if hint:
            print(f"  {hint}")
        print("  Copy .env.example to .env and fill it in.\n")
        sys.exit(1)
    return val


def _looks_like_anon(key):
    """True if `key` is a publishable/anon key rather than a secret one."""
    if key.startswith("sb_publishable_"):
        return True
    try:
        payload = key.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload)).get("role") == "anon"
    except Exception:
        return False


def supabase_config(require_service_role=True):
    """Return (url, key). Raises a clear error instead of falling back to a
    key baked into the source, which is what the old version did."""
    url = _required(
        "SUPABASE_URL",
        "Supabase dashboard -> Settings -> API -> Project URL",
    ).rstrip("/")

    if require_service_role:
        key = _required(
            "SUPABASE_SERVICE_ROLE_KEY",
            "Supabase dashboard -> Settings -> API -> service_role (secret)",
        )
        if _looks_like_anon(key):
            print(
                "\nSUPABASE_SERVICE_ROLE_KEY looks like an anon/publishable key."
                "\nInserts and deletes will be silently blocked by Row Level Security.\n"
            )
            sys.exit(1)
    else:
        key = _required(
            "SUPABASE_ANON_KEY",
            "Supabase dashboard -> Settings -> API -> anon (public)",
        )
    return url, key


def env_path(name, fallback=""):
    """Configurable filesystem path, so hardcoded E:\\ drives stay overridable."""
    return (os.environ.get(name) or fallback).strip()


# ── date normalisation ─────────────────────────────────────────────────────

# Day-first order matters: "07/08/2026" is genuinely ambiguous, and this
# pipeline's source data is day/month/year. Change the order here if your CRM
# ever exports US-style month-first dates.
_DATE_FORMATS = (
    "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y",
    "%Y/%m/%d", "%d.%m.%Y", "%d-%b-%Y", "%b %d, %Y",
)


def clean_date(value):
    """Normalise anything Excel/pandas hands us into an ISO 'YYYY-MM-DD' string.

    This is the fix for the duplicate-row bug: pandas returns a Timestamp for
    date-formatted cells, and str() on that yields '2026-07-01 00:00:00'.
    Postgres stores it as '2026-07-01', so every comparison between a freshly
    read row and an existing DB row failed, and the de-duplication step let
    the same visit through on every single re-upload.

    Returns None for blanks so the column stays NULL rather than the string
    "NaT".
    """
    if value is None:
        return None

    # NaT must be tested BEFORE the datetime check: pandas' NaTType subclasses
    # datetime, so isinstance(NaT, datetime) is True and calling .strftime on
    # it raises instead of returning None.
    try:
        import pandas as pd
        if pd.isna(value):
            return None
    except (TypeError, ValueError, ImportError):
        pass  # pd.isna raises on arrays/lists; those fall through to str()

    # pandas Timestamp / datetime / date all expose .strftime
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")

    s = str(value).strip()
    if not s or s.lower() in {"nan", "nat", "none", "null"}:
        return None

    # Already ISO, possibly with a time component tacked on.
    head = s[:10]
    try:
        datetime.strptime(head, "%Y-%m-%d")
        return head
    except ValueError:
        pass

    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s.split(" ")[0], fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue

    # Excel serial number (days since 1899-12-30)
    try:
        serial = float(s)
        if 1 < serial < 100000:
            from datetime import timedelta
            return (datetime(1899, 12, 30) + timedelta(days=serial)).strftime("%Y-%m-%d")
    except ValueError:
        pass

    return s  # unrecognised: pass through so the DB error is visible, not silent


def clean_time(value):
    """Normalise a time cell to 'HH:MM' or 'HH:MM:SS'."""
    if value is None:
        return None
    try:
        import pandas as pd
        if pd.isna(value):
            return None
    except (TypeError, ValueError, ImportError):
        pass
    if isinstance(value, datetime):
        return value.strftime("%H:%M:%S")
    s = str(value).strip()
    if not s or s.lower() in {"nan", "nat", "none"}:
        return None
    # '2026-07-01 09:30:00' -> '09:30:00'
    if " " in s and "-" in s.split(" ")[0]:
        s = s.split(" ", 1)[1]
    return s


# ── file helpers ───────────────────────────────────────────────────────────

def list_workbooks(folder):
    """All real .xlsx files in `folder`, excluding Excel's '~$' lock files.

    Excel creates '~$Report.xlsx' the moment somebody opens a workbook. The
    old code globbed '*.xlsx' and then took the newest by mtime, so an open
    workbook made the uploader pick the lock file — which is not a valid
    workbook and blows up mid-run.
    """
    p = Path(folder)
    if not p.is_dir():
        return []
    return [
        f for f in p.glob("*.xlsx")
        if not f.name.startswith("~$") and not f.name.startswith(".") and f.is_file()
    ]


def wait_until_stable(path, checks=3, interval=1.0, timeout=60):
    """Block until a file stops growing, so we never read a half-written export."""
    import time
    path = Path(path)
    last, stable, waited = -1, 0, 0.0
    while waited < timeout:
        try:
            size = path.stat().st_size
        except OSError:
            size = -1
        if size == last and size >= 0:
            stable += 1
            if stable >= checks:
                return True
        else:
            stable = 0
        last = size
        time.sleep(interval)
        waited += interval
    return False


# ── time parsing ───────────────────────────────────────────────────────────

_TIME_FORMATS = ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p", "%I:%M%p", "%I%p")


def parse_time_minutes(value):
    """Parse a clock time into minutes since midnight, or None.

    Each side is parsed independently. The previous implementation looped
    over formats and required BOTH times to match the SAME format, so
    comparing '09:30' with '09:45:00' fell through every branch and returned
    a sentinel meaning "far apart" — silently dropping genuine matches in the
    coaching-day accompaniment check.
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.hour * 60 + value.minute
    s = str(value).strip()
    if not s or s.lower() in {"nan", "nat", "none"}:
        return None
    # '2026-07-01 09:30:00' -> '09:30:00'
    if " " in s and "-" in s.split(" ")[0]:
        s = s.split(" ", 1)[1].strip()
    s = s.upper().replace(".", "")
    for fmt in _TIME_FORMATS:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.hour * 60 + dt.minute
        except ValueError:
            continue
    return None


def minutes_between(t1, t2, unparseable=999):
    """Absolute minutes between two clock times.

    Returns `unparseable` (a deliberately large sentinel) if either side
    can't be read, so a missing time reads as "far apart" rather than
    accidentally counting as a match.
    """
    a = parse_time_minutes(t1)
    b = parse_time_minutes(t2)
    if a is None or b is None:
        return unparseable
    return abs(b - a)


def sort_times(times):
    """Chronological sort of clock-time strings.

    sorted() on raw strings is LEXICOGRAPHIC: ['9:30','10:15'] sorts to
    ['10:15','9:30'], and '2:00 PM' sorts before '9:30 AM'. Shift start/end
    were read off sorted()[0] and [-1], so any unpadded or 12-hour time made
    the duration and the average start time wrong.
    """
    return sorted(times, key=lambda t: (parse_time_minutes(t) is None,
                                        parse_time_minutes(t) or 0,
                                        str(t)))


def fmt_hm(total_minutes):
    """Format minutes as 'h:mm'. Negative or invalid input clamps to '0:00'.

    The old version used floor division on negatives, so -30 rendered as
    '-1:30' instead of '0:00'.
    """
    try:
        total = float(total_minutes)
    except (TypeError, ValueError):
        return "0:00"
    if total != total or total < 0:  # NaN or negative
        return "0:00"
    return f"{int(total // 60)}:{int(total % 60):02d}"
