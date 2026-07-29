/**
 * generate_hierarchy_sql.js — produce a transactional SQL file that replaces
 * the `hierarchy` table, for pasting into the Supabase SQL editor.
 *
 * FIXED:
 *  - Service-role key no longer hardcoded; read from .env via config.js.
 *    (It is only used here to look up team IDs — a read.)
 *  - Input and output paths are configurable via .env or argv.
 *  - Duplicate employee codes are detected and reported instead of producing
 *    a SQL file that violates a unique constraint halfway through.
 *  - Aborts on an empty spreadsheet rather than emitting a file whose only
 *    effect is "DELETE FROM hierarchy".
 *
 * The generated SQL is wrapped in BEGIN/COMMIT, so unlike a chunked REST
 * delete-then-insert it either fully applies or fully rolls back.
 */

const fs = require('fs');
const path = require('path');
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

/** Quote a value for SQL, doubling embedded single quotes. */
const escSql = v => {
    if (v === null || v === undefined) return 'NULL';
    return "'" + String(v).replace(/'/g, "''") + "'";
};

async function generateSQL() {
    const args = process.argv.slice(2).filter(a => !a.startsWith('--'));
    const filePath = args[0] || envPath('HIERARCHY_FILE', 'E:\\crm extractor\\hierarchy\\hierarchy_export.xlsx');
    const outPath = args[1] || envPath('HIERARCHY_SQL_OUT', 'E:\\crm extractor\\hierarchy\\hierarchy_replace.sql');

    console.log(`Reading ${filePath}...`);
    if (!fs.existsSync(filePath)) {
        console.error(`File not found: ${filePath}`);
        console.error('Set HIERARCHY_FILE in .env or pass the path as the first argument.');
        process.exitCode = 1;
        return;
    }

    let data;
    try {
        const wb = xlsx.readFile(filePath);
        data = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });
    } catch (err) {
        console.error(`Could not read the workbook: ${err.message}`);
        process.exitCode = 1;
        return;
    }

    if (!data.length) {
        console.error('No data found in the Excel file — refusing to generate a delete-only script.');
        process.exitCode = 1;
        return;
    }

    console.log(`Found ${data.length} records in Excel.`);
    console.log(`Excel columns: ${Object.keys(data[0]).join(', ')}`);

    // --- team name -> id --------------------------------------------------
    console.log('\nFetching teams table from Supabase...');
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

    // --- map rows ---------------------------------------------------------
    const seen = new Set();
    const duplicates = [];
    const records = [];

    for (const row of data) {
        const empCode = cleanCode(row['employee_code'] || row['Employee Code'] || row['Employee_Code']);
        if (!empCode) continue;
        if (seen.has(empCode)) {
            duplicates.push(empCode);
            continue;
        }
        seen.add(empCode);

        const teamName = cleanStr(row['team'] || row['Team'] || row['team_name'] || row['Team Name']);
        records.push({
            employee_code: empCode,
            employee_name: cleanStr(row['employee_name'] || row['Employee Name'] || row['Employee_Name']),
            role: cleanStr(row['role'] || row['Role']),
            supervisor_name: cleanStr(row['supervisor_name'] || row['Supervisor Name'] || row['Supervisor_Name']),
            area_manager_name: cleanStr(row['area_manager_name'] || row['Area Manager Name'] || row['Area_Manager_Name']),
            team_id: teamName ? (teamNameToId[teamName.toLowerCase()] ?? null) : null,
            _team_name: teamName,
        });
    }

    if (duplicates.length) {
        console.warn(`\nWARNING: ${duplicates.length} duplicate employee_code(s); kept the first of each:`);
        [...new Set(duplicates)].forEach(c => console.warn(`   - ${c}`));
    }

    const unmappedTeams = [...new Set(records.filter(r => r._team_name && !r.team_id).map(r => r._team_name))];
    if (unmappedTeams.length) {
        console.warn(`\nWARNING: these team names don't match any team in Supabase:`);
        unmappedTeams.forEach(t => console.warn(`   - "${t}"`));
        console.warn('   They will be written with team_id = NULL.');
    }

    if (!records.length) {
        console.error('\nNo rows had an employee_code — nothing to generate.');
        process.exitCode = 1;
        return;
    }

    // --- build SQL --------------------------------------------------------
    const valueRows = records.map(r =>
        `  (${escSql(r.employee_code)}, ${escSql(r.employee_name)}, ${escSql(r.role)}, ` +
        `${escSql(r.supervisor_name)}, ${escSql(r.area_manager_name)}, ` +
        `${r.team_id !== null ? r.team_id : 'NULL'})`
    );

    const sql =
        '-- ===================================================\n' +
        '-- Hierarchy replacement SQL\n' +
        `-- Generated: ${new Date().toISOString()}\n` +
        `-- Source: ${path.basename(filePath)} (${records.length} records)\n` +
        '--\n' +
        '-- Wrapped in a transaction: if any row fails, the whole thing rolls\n' +
        '-- back and the existing hierarchy is left intact.\n' +
        '-- ===================================================\n\n' +
        'BEGIN;\n\n' +
        '-- Step 1: Delete all existing hierarchy rows\n' +
        'DELETE FROM hierarchy;\n\n' +
        '-- Step 2: Insert new hierarchy data\n' +
        'INSERT INTO hierarchy (employee_code, employee_name, role, supervisor_name, area_manager_name, team_id)\nVALUES\n' +
        valueRows.join(',\n') + ';\n\n' +
        '-- Step 3: Sanity check — rolls back if the row count looks wrong\n' +
        'DO $$\nDECLARE n integer;\nBEGIN\n' +
        '  SELECT count(*) INTO n FROM hierarchy;\n' +
        `  IF n <> ${records.length} THEN\n` +
        `    RAISE EXCEPTION 'Expected ${records.length} rows, found %', n;\n` +
        '  END IF;\nEND $$;\n\n' +
        'COMMIT;\n';

    try {
        fs.mkdirSync(path.dirname(outPath), { recursive: true });
        fs.writeFileSync(outPath, sql, 'utf8');
    } catch (err) {
        console.error(`Could not write ${outPath}: ${err.message}`);
        process.exitCode = 1;
        return;
    }

    console.log(`\nSQL file written to: ${outPath}`);
    console.log(`   ${records.length} INSERT rows generated.`);
    console.log(`\nCopy the contents of that file into the Supabase SQL Editor and run it.`);
}

generateSQL().catch(err => {
    console.error('Unexpected failure:', err);
    process.exitCode = 1;
});
