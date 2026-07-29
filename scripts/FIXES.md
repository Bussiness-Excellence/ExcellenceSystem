# Excellence CRM — what was fixed

---

## READ THIS FIRST: rotate your keys

Five files contained live credentials in plaintext:

| File | Key that was exposed |
|---|---|
| `emergency_upload.js` | `service_role` JWT |
| `generate_hierarchy_sql.js` | `service_role` JWT |
| `update_hierarchy.js` | `service_role` JWT |
| `pulpoplus_upload_to_supabase.py` | `sb_secret_…` (in `_supabase_config()` **and** in the docstring) |

| `pulpoplus_extractor_summary.py` | **PulpoPlus CRM login** — URL, username and password |

`push_hierarchy.py` and `inspect_data.py` also carried a publishable key.

That last one is the worst of the set: it is the login to the CRM system
itself, in plaintext, in a repo you push to GitHub. Change that password in
PulpoPlus, not just in the file.

The `service_role` key bypasses Row Level Security completely — it is the
database equivalent of a root password. You push this project to GitHub, so
assume both secrets are public.

**Do this before anything else:**

1. Supabase dashboard → Settings → API → rotate the `service_role` key
   (and the publishable key).
1a. Change the PulpoPlus CRM account password.
2. Copy `.env.example` to `.env` and paste the **new** keys in.
3. Confirm `.env` is ignored: `git check-ignore -v .env` should print a match.
4. Rotating does not erase the old key from git history. To purge it, use
   [`git filter-repo`](https://github.com/newren/git-filter-repo) or BFG, then
   force-push. Rotating is the part that actually stops the bleeding; history
   cleanup is housekeeping.
5. Check Supabase → Logs for unexpected activity while the key was public.

Nothing in this bundle contains a credential. Everything reads from `.env`.

---

## The bug that was corrupting your dashboard numbers

Three separate defects stacked up to produce duplicated visits and coaching
days. All three are fixed.

### 1. Dates never matched, so de-duplication never worked

`pulpoplus_upload_to_supabase.py` built the date column with
`clean_str(r.get("date"))`. Pandas returns a `Timestamp` for date-formatted
cells, and `str()` on that yields `2026-07-01 00:00:00`. Postgres stores a
`date` column as `2026-07-01`.

The `--append` path de-duplicates by comparing
`(user, visit_date, visit_time, acc_id)` against rows already in the database.
Those tuples could never match, so **every single re-upload re-inserted every
row.** Demonstrated:

```
old clean_str  -> '2026-07-01 00:00:00'   matches DB?  False
new clean_date -> '2026-07-01'            matches DB?  True
```

`emergency_upload.js` had the same defect via `String(dateValue)`, producing
`Wed Jul 01 2026 00:00:00 GMT+0200 (…)`.

Fixed with a shared `clean_date()` / `cleanDate()` that handles Timestamps,
`NaT`, ISO strings, `dd/mm/yyyy`, and Excel serial numbers. Note: `07/08/2026`
is genuinely ambiguous — the parser assumes **day first**, matching your source
data. The order is documented at the top of `_DATE_FORMATS` if that ever changes.

### 2. `supabase_select` fetched at most 1000 rows

PostgREST caps every response at 1000 rows. The old helper issued one
un-paginated request, so the de-duplication check only ever saw the first
1000 existing rows. Past that — which every real month exceeds — re-running
an upload duplicated everything beyond row 1000, even once the dates matched.

Now paginates with `Range` headers until the table is exhausted.

### 3. Batch labels didn't line up

`pulpoplus_auto_upload.py` writes rows tagged `recent_July 2026 - Focus 1`
(batch + `_` + filename), but `emergency_upload.js` cleared with
`.eq('upload_batch', 'recent')`. Those never match, so the "clear old data"
step silently deleted nothing and the restore stacked a fresh copy on top.

Now clears with `upload_batch.eq.X OR upload_batch.like.X\_%`, covering both
shapes.

**After deploying:** run `node check_visits.js "<name>" <YYYY-MM-DD>` — it now
flags duplicated rows. Existing duplicates in the database are not removed
automatically; clear the affected batch and re-upload once.

---

## Everything else

### `upload_now.bat` — could never have run

The Python was embedded inline after `python -c "`. cmd.exe ends a command at
the newline, so the quote was never closed and every following line was
executed as its own shell command. Moved to a real `upload_now.py`; the
`.bat` now just calls it and checks the exit code.

### `run_auto_uploader.bat` — unescaped `&`

`echo Starting … Watcher & Uploader…` made cmd try to *execute* `Uploader…`.
Escaped to `^&`. The echoed paths (`E:\periods\…`) also didn't match the real
defaults; they now point at the `.env` settings. All three `.bat` files were
rewritten with CRLF line endings and `errorlevel` checks.

### Excel lock files were being uploaded

`Path(folder).glob("*.xlsx")` matches `~$Report.xlsx` — the lock file Excel
creates the moment anyone opens a workbook. Since it's always the newest file,
`max(mtime)` picked it every time, and it isn't a readable workbook.
`list_workbooks()` now filters `~$` and `.` prefixes.

Also added `wait_until_stable()`, so the watcher never reads a workbook that is
still being written or synced.

### Watcher debounce dropped events

If a file changed while an upload was running, the event was discarded. The
handler now records that work arrived mid-run and repeats, guarded by a lock.

### `update_hierarchy.js` — destructive replace with no transaction

Old flow: `DELETE` every row, then `INSERT` in chunks. A failure on chunk 3 of
5 left a half-empty org chart with nothing to roll back to, and broke every
foreign key pointing at those rows.

Now upserts on `employee_code`, and only deletes rows genuinely absent from the
spreadsheet — after all inserts have succeeded. Aborts without touching
anything if the spreadsheet is empty or unreadable.

**One-time setup required:**

```sql
create unique index if not exists hierarchy_employee_code_key
  on hierarchy (employee_code);
```

Pass `--replace` for the old delete-everything behaviour.

### `push_hierarchy.py` — anon key doing privileged deletes

It used a publishable key to `DELETE FROM hierarchy`. Either that silently
failed (and the script carried on inserting on top of the old rows), or it
succeeded — which would mean `hierarchy` has no RLS policy and anyone holding
your public key can wipe it. **Worth checking which.** Now requires the service
role key and stops on the first failed batch instead of pressing on.

### `generate_hierarchy_sql.js`

Credentials removed; paths configurable. Added a duplicate-`employee_code`
check and a `DO $$` row-count assertion inside the transaction, so a wrong
count rolls the whole thing back.

### `create_monorepo.py` — copied `.env` into the repo

The copy list explicitly included `.env`, walking your keys straight into the
folder you push to GitHub. It now copies `.env.example` instead, writes a
`.gitignore` covering both halves of the tree, and sweeps the output for stray
`.env*` files afterwards, exiting non-zero if it finds any.

### Theme scripts lied about succeeding

`fix_css.js`, `alvent_theme.js`, and `salient_theme.js` ran a chain of regex
replacements, then wrote the file and printed "applied successfully!"
regardless. `String.replace()` returns the input unchanged when nothing
matches — so a regex that had drifted out of sync failed **completely
silently**.

Now each substitution is labelled and reports `applied` or `NO MATCH`, the
file is backed up to `.bak` before writing, nothing is written if nothing
changed, and the script exits non-zero if any rule missed.

`fix_css.js` also had three byte-identical patterns in its replacement list
(header / ctrl-bar / tabs). With the `/g` flag the first consumed every
occurrence, so entries 2 and 3 were dead code. Collapsed into one.

Re-running a theme script no longer stacks duplicate `backdrop-filter`
declarations.

### Smaller items

- `supabase_delete_all()` filtered on `id=gt.0`, which matches nothing on a
  uuid primary key — a "full resync" silently left old rows in place. Now uses
  `id=not.is.null`.
- Summary de-duplication keyed only on `user_name`; two reps sharing a display
  name collided. Now keyed on `(user_name, employee_code)`.
- `emergency_upload.js` logged insert errors and carried on, leaving partial
  data with a success message at the end. Now stops and exits non-zero.
- `check_visits.js` and `inspect_excel.js` had names, dates and file paths
  hardcoded. Now arguments.
- `inspect_excel.js` checked four different date string formats by hand —
  itself a symptom of dates never being normalised. Now one normalised check.
- `inspect_data.py` uses the anon key deliberately: it only reads, so it
  shouldn't hold a key that can delete.
- Both config loaders reject a publishable key pasted into the service-role
  slot, instead of letting writes fail confusingly at the RLS layer.

---

## Second pass: the four extractor / rebuild files

These were unreviewed in the first bundle. Auditing them turned up the
credential leak above plus six more defects.

### Time comparison silently failed on mixed formats

`minutes_between()` looped over format strings and required **both** times to
parse with the **same** one:

```
'09:30'    vs '09:45'       -> 15    ok
'09:30'    vs '09:45:00'    -> 999   WRONG
'09:30'    vs '10:00 AM'    -> 999   WRONG
```

999 is the "too far apart" sentinel, so genuinely matching visits were read as
unrelated. This feeds the coaching-day accompaniment check directly — real
coaching days were being dropped whenever the export mixed time formats.

The copy in `pulpoplus_extractor_summary.py` was worse: a bare `strptime` with
`"%H:%M"` and **no** try/except, so one `"9:30 AM"` value raised `ValueError`
and killed the whole extraction run.

Both now use a shared helper that parses each side independently.

### Shift durations and start times used a lexicographic sort

`sorted(times)` on raw strings is alphabetical, not chronological:

```
['9:30', '10:15', '14:00']         -> ['10:15', '14:00', '9:30']
['9:30 AM', '11:00 AM', '2:00 PM'] -> ['11:00 AM', '2:00 PM', '9:30 AM']
```

The code takes `sorted(times)[0]` and `[-1]` as the shift's start and end. With
any unpadded or 12-hour value, it picked the wrong endpoints — so **Avg AM/PM
Shift Duration and Avg AM Starting Time were wrong**. Zero-padded 24-hour times
happen to sort correctly, which is why this held up as long as it did.

Replaced with a `sort_times()` that sorts by parsed minutes.

### Coaching accompaniment failed open on missing timestamps

```python
if not m_time or not r["time"] or minutes_between(...) <= TOLERANCE:
    matched_managers.add(m_user)
```

A missing timestamp on either side counted as a **confirmed** time match, so
gaps in the export inflated coaching days. Same account + date + shift is still
treated as evidence, but those matches are now counted and reported separately
rather than being indistinguishable from verified ones.

### `_rows_of()` discarded the first data row of every section

```python
return [r for r in section_html.split("<tr") if "<td" in r][1:]
```

The `[1:]` assumed the first row is a header. But the filter keeps only rows
containing `<td>`, and header cells are normally `<th>` — so the header was
already excluded and the `[1:]` **silently dropped the first real data row of
every section**. Now the first row is only dropped if it actually looks like a
header (`<th>` present, or the cell text matches column labels).

### Malformed-row skips were invisible

Every section parser counts rows it skips for having an unexpected column
count, but only printed that count under `--debug`. A column change in the CRM
export would drop records with no visible sign. Skips are now always reported.

### Manager detection is a load-bearing guess

`identify_managers()` classifies anyone whose records span more than one
territory as a manager. Anyone flagged is excluded from rep visit totals and
from coaching de-duplication, so a misclassification silently reshapes Total
Visits, AM/PM Calls and Coaching Days.

Two failure modes, both confirmed with a test:

- a rep who **transfers territory mid-period** is promoted to manager and drops
  out of the rep numbers entirely;
- a manager who **only worked one territory** that period is never detected, so
  their coaching days are never counted.

You already have an org chart with explicit roles. `identify_managers()` now
takes the hierarchy, treats its roles as authoritative for anyone it knows
about, falls back to the heuristic for anyone missing from it, and **prints
every disagreement**. Verified:

```
heuristic only        -> ['Mgr B', 'Rep A']      Rep A wrong, Mgr C missed
with hierarchy check  -> ['Mgr B', 'Mgr C']      both corrected, both reported
```

Expect your coaching-day and visit numbers to move after this. That is the
point — but check the reported disagreements against reality before trusting
the new figures.

---

## Still needs a decision from you

**The two engines define a coaching day differently.** Same input, different
answers:

| | `pulpoplus_rebuild_summary.py` | `pulpoplus_extractor_summary.py` |
|---|---|---|
| Rule | manager accompanied **≥80%** of the rep's calls in *both* shifts | **≥1** matched visit in *both* shifts |

The second is far easier to satisfy, so it reports more coaching days. I have
not changed either — which one is correct is a business question, not a coding
one. But whichever path a given month went through is currently determining the
number, and that shouldn't be true. Pick one rule and I can align the other.

`pulpoplus_hierarchy.py` and `generate_docs.py` were audited and left alone —
the narrow `except Exception: pass` blocks in the hierarchy parser are guarded
integer conversions and are fine as they stand.

What I have **not** verified, because it needs real data to check against: the
AM/PM shift category mappings, the 90-minute proximity tolerance, the 80%
threshold, the half-day (0.5) counting rule for activities and office work, and
the territory backfill's majority-vote fallback. Those are all judgement calls
encoded as constants; they are consistent, but whether they match how your
business actually defines these things I can't tell from the code alone.

## Setup

```bash
cp .env.example .env      # then fill in the NEW keys + PULPO_USERNAME/PASSWORD
npm install               # @supabase/supabase-js, dotenv, xlsx
pip install -r requirements.txt --break-system-packages
```

Then the one-time SQL index for `update_hierarchy.js` (above).

## Suggested order

1. Rotate the Supabase keys **and change the CRM password**, fill in `.env`,
   verify `.gitignore`.
2. Create the unique index on `hierarchy(employee_code)`.
3. Check whether the anon key can delete from `hierarchy` — if it can, add an
   RLS policy.
4. Run `node check_visits.js "<name>" <date>` to see current duplicates.
5. Clear the affected batch and re-upload with the fixed scripts.
6. Confirm the numbers match the source workbook.
7. Review the manager-mismatch warnings the rebuild now prints — those change
   the totals.
8. Decide which coaching-day rule is correct and align the two engines.
