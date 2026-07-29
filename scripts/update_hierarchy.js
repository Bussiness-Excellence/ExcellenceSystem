/**
 * update_hierarchy.js — sync hierarchy_export.xlsx into the Supabase
 * `hierarchy` table.
 *
 * FIXED:
 *  - Service-role key is no longer hardcoded; it comes from .env via config.js.
 *  - No longer deletes every row before inserting. The old flow was
 *    DELETE ALL -> INSERT in chunks with no transaction, so a failure on
 *    chunk 3 of 5 left a half-empty org chart with nothing to roll back to,
 *    and broke every foreign key pointing at those rows.
 *    It now upserts on employee_code and only removes rows genuinely gone
 *    from the spreadsheet — and only after all inserts have succeeded.
 *  - Aborts instead of wiping anything if the spreadsheet is empty/unreadable.
 *  - The file path is configurable via HIERARCHY_FILE or argv.
 *
 * NOTE: upserting on employee_code requires a unique index on that column:
 *     create unique index if not exists hierarchy_employee_code_key
 *       on hierarchy (employee_code);
 * Run that once in the SQL editor. Pass --replace to fall back to the old
 * delete-everything behaviour if you really want it.
 */

const fs = require('fs');
const xlsx = require('xlsx');
const { getServiceClient, envPath } = require('./config');

const supabase = getServiceClient();

const cleanStr = v =>
    (v === null || v === undefined || v === '' || v === 'NaN') ? null : String(v).trim();

const cleanCode = v => {
    let s = cleanStr(v);
    if (!s) return null;
    if (s.endsWith('.0')) s = s.slice(0, -2);
    return s;
};

const argPath = process.argv.slice(2).find(a => !a.startsWith('--'));
const REPLACE_MODE = process.argv.includes('--replace');

async function updateHierarchy() {
    const filePath =
        argPath || envPath('HIERARCHY_FILE', 'E:\\crm extractor\\hierarchy\\hierarchy_export.xlsx');

    console.log(`Reading ${filePath}...`);
    if (!fs.existsSync(filePath)) {
        console.error(`File not found: ${filePath}`);
        console.error(`Set HIERARCHY_FILE in .env or pass the path as an argument.`);
        process.exitCode = 1;
        return;
    }

    let data;
    try {
        const wb = xlsx.readFile(filePath);
        const sheetName = wb.SheetNames[0];
        data = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });
    } catch (err) {
        console.error(`Could not read the workbook: ${err.message}`);
        process.exitCode = 1;
        return;
    }

    if (!data.length) {
        console.error('No data found in the Excel file. Aborting without touching the table.');
        process.exitCode = 1;
        return;
    }

    console.log(`Found ${data.length} records in Excel.`);
    console.log(`Excel columns: ${Object.keys(data[0]).join(', ')}`);

    // --- Step 1: map team names -> team IDs ---------------------------------
    console.log('\nFetching teams table...');
    const { data: teams, error: teamsErr } = await supabase.from('teams').select('id,name');
    if (teamsErr) {
        console.error('Error fetching teams:', teamsErr.message);
        process.exitCode = 1;
        return;
    }
    const teamNameToId = {};
    for (const t of teams) {
        if (t.name) teamNameToId[t.name.toLowerCase()] = t.id;
    }
    console.log(`Found ${teams.length} teams: ${teams.map(t => `${t.name} (id=${t.id})`).join(', ')}`);

    // --- Step 2: map spreadsheet rows to DB shape ---------------------------
    const seen = new Set();
    const duplicates = [];
    const cleanRecords = [];

    for (const row of data) {
        const empCode = cleanCode(row['employee_code'] || row['Employee Code'] || row['Employee_Code']);
        if (!empCode) continue;

        // Upsert needs one row per conflict key, or Postgres rejects the whole
        // statement with "cannot affect row a second time".
        if (seen.has(empCode)) {
            duplicates.push(empCode);
            continue;
        }
        seen.add(empCode);

        const teamName = cleanStr(row['team'] || row['Team'] || row['team_name'] || row['Team Name']);
        const teamId = teamName ? (teamNameToId[teamName.toLowerCase()] ?? null) : null;

        cleanRecords.push({
            employee_code: empCode,
            employee_name: cleanStr(row['employee_name'] || row['Employee Name'] || row['Employee_Name']),
            role: cleanStr(row['role'] || row['Role']),
            supervisor_name: cleanStr(row['supervisor_name'] || row['Supervisor Name'] || row['Supervisor_Name']),
            area_manager_name: cleanStr(row['area_manager_name'] || row['Area Manager Name'] || row['Area_Manager_Name']),
            team_id: teamId,
            _team_name: teamName,
        });
    }

    if (duplicates.length) {
        console.warn(`\nWARNING: ${duplicates.length} duplicate employee_code(s) in the spreadsheet; kept the first of each:`);
        [...new Set(duplicates)].forEach(c => console.warn(`   - ${c}`));
    }

    const unmappedTeams = [...new Set(
        cleanRecords.filter(r => r._team_name && !r.team_id).map(r => r._team_name)
    )];
    if (unmappedTeams.length) {
        console.warn(`\nWARNING: these team names don't match any team in Supabase:`);
        unmappedTeams.forEach(t => console.warn(`   - "${t}"`));
        console.warn(`   Those people will be stored with team_id = null.`);
    }

    const payload = cleanRecords.map(({ _team_name, ...rest }) => rest);
    if (!payload.length) {
        console.error('\nNo rows had an employee_code. Aborting without touching the table.');
        process.exitCode = 1;
        return;
    }

    // --- Step 3 (optional): legacy full replace -----------------------------
    if (REPLACE_MODE) {
        console.log('\n--replace given: deleting all existing hierarchy rows first...');
        const { error: delErr } = await supabase.from('hierarchy').delete().not('id', 'is', null);
        if (delErr) {
            console.error('Delete failed, nothing was changed:', delErr.message);
            process.exitCode = 1;
            return;
        }
    }

    // --- Step 4: upsert -----------------------------------------------------
    console.log(`\nUpserting ${payload.length} records (chunks of 500)...`);
    let totalUpserted = 0;
    let failed = false;

    for (let i = 0; i < payload.length; i += 500) {
        const chunk = payload.slice(i, i + 500);
        const { error } = REPLACE_MODE
            ? await supabase.from('hierarchy').insert(chunk)
            : await supabase.from('hierarchy').upsert(chunk, { onConflict: 'employee_code' });

        if (error) {
            failed = true;
            console.error(`  Error on rows ${i}-${i + chunk.length}: ${error.message}`);
            console.error(`  First record in the failing chunk:`, JSON.stringify(chunk[0], null, 2));
            if (/no unique|constraint matching/i.test(error.message)) {
                console.error(
                    '\n  The upsert needs a unique index on employee_code. Run once:\n' +
                    '    create unique index if not exists hierarchy_employee_code_key\n' +
                    '      on hierarchy (employee_code);\n'
                );
            }
            break;  // stop early rather than continuing to half-write
        }
        totalUpserted += chunk.length;
        console.log(`  Upserted rows ${i} to ${i + chunk.length}`);
    }

    if (failed) {
        console.error(`\nStopped after ${totalUpserted} rows. Existing data was NOT deleted, so the table is still usable.`);
        process.exitCode = 1;
        return;
    }

    // --- Step 5: remove people no longer in the spreadsheet -----------------
    // Only runs once every insert has succeeded, so a mid-run failure can
    // never leave the org chart truncated.
    if (!REPLACE_MODE) {
        const codes = new Set(payload.map(r => r.employee_code));
        const { data: existing, error: exErr } = await supabase
            .from('hierarchy')
            .select('id,employee_code,employee_name');

        if (exErr) {
            console.warn(`\nCould not check for stale rows: ${exErr.message}`);
        } else {
            const stale = existing.filter(r => !codes.has(r.employee_code));
            if (stale.length) {
                console.log(`\nRemoving ${stale.length} row(s) no longer present in the spreadsheet:`);
                stale.slice(0, 20).forEach(r =>
                    console.log(`   - ${r.employee_name || '(no name)'} [${r.employee_code}]`));
                if (stale.length > 20) console.log(`   ... and ${stale.length - 20} more`);

                for (let i = 0; i < stale.length; i += 500) {
                    const ids = stale.slice(i, i + 500).map(r => r.id);
                    const { error: dErr } = await supabase.from('hierarchy').delete().in('id', ids);
                    if (dErr) console.error(`  Could not delete stale rows: ${dErr.message}`);
                }
            } else {
                console.log('\nNo stale rows to remove.');
            }
        }
    }

    console.log(`\nDone. ${totalUpserted} / ${payload.length} records synced to the hierarchy table.`);
}

updateHierarchy().catch(err => {
    console.error('Unexpected failure:', err);
    process.exitCode = 1;
});
