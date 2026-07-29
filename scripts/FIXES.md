# Excellence CRM — what was fixed

---

## READ THIS FIRST: rotate your keys

Four files contained live Supabase credentials in plaintext:

| File | Key that was exposed |
|---|---|
| `emergency_upload.js` | `service_role` JWT |
| `generate_hierarchy_sql.js` | `service_role` JWT |
| `update_hierarchy.js` | `service_role` JWT |
| `pulpoplus_upload_to_supabase.py` | `sb_secret_…` (in `_supabase_config()` **and** in the docstring) |

`push_hierarchy.py` and `inspect_data.py` also carried a publishable key.

The `service_role` key bypasses Row Level Security completely — it is the
database equivalent of a root password. You push this project to GitHub, so
assume both secrets are public.

**Do this before anything else:**

1. Supabase dashboard → Settings → API → rotate the `service_role` key
   (and the publishable key).
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

## Not changed

`pulpoplus_extract_visits.py`, `pulpoplus_extractor_summary.py`,
`pulpoplus_hierarchy.py`, `pulpoplus_rebuild_summary.py` and
`generate_docs.py` are unmodified — no credentials and no bugs found in the
areas reviewed. Their business logic (dedup rules, shift calculations,
coaching-day matching) wasn't audited; that would need sample data to verify
against.

---

## Setup

```bash
cp .env.example .env      # then fill in the NEW keys
npm install               # @supabase/supabase-js, dotenv, xlsx
pip install -r requirements.txt --break-system-packages
```

Then the one-time SQL index for `update_hierarchy.js` (above).

## Suggested order

1. Rotate keys, fill in `.env`, verify `.gitignore`.
2. Create the unique index on `hierarchy(employee_code)`.
3. Check whether the anon key can delete from `hierarchy` — if it can, add an
   RLS policy.
4. Run `node check_visits.js "<name>" <date>` to see current duplicates.
5. Clear the affected batch and re-upload with the fixed scripts.
6. Confirm the numbers match the source workbook.
