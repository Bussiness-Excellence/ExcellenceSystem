const fs = require('fs');
const xlsx = require('xlsx');
const { createClient } = require('@supabase/supabase-js');

const supabaseUrl = 'https://xxbfwvlqixnmonxytdxq.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh4YmZ3dmxxaXhubW9ueHl0ZHhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4Mjc1NjE2NSwiZXhwIjoyMDk4MzMyMTY1fQ.PSk6RyFmg_OFTcCtYO74AeJj6wT4FGZS2K2JT9GEJ_A';
const supabase = createClient(supabaseUrl, supabaseKey);

const cleanStr = v => (v === null || v === undefined || v === '' || v === 'NaN') ? null : String(v).trim();
const cleanCode = v => {
    let s = cleanStr(v);
    if (!s) return null;
    if (s.endsWith('.0')) s = s.slice(0, -2);
    return s;
};

const escSql = v => {
    if (v === null || v === undefined) return 'NULL';
    return "'" + String(v).replace(/'/g, "''") + "'";
};

async function generateSQL() {
    const filePath = 'E:\\crm extractor\\hierarchy\\hierarchy_export.xlsx';
    console.log(`Reading ${filePath}...`);

    if (!fs.existsSync(filePath)) {
        console.error(`File not found: ${filePath}`);
        return;
    }

    const wb = xlsx.readFile(filePath);
    const sheetName = wb.SheetNames[0];
    const data = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });

    if (data.length === 0) {
        console.log("No data found in the Excel file.");
        return;
    }

    console.log(`Found ${data.length} records in Excel.`);
    console.log(`Excel columns: ${Object.keys(data[0]).join(', ')}`);

    // --- Fetch the teams table to map team names -> team IDs ---
    console.log("\nFetching teams table from Supabase...");
    const { data: teams, error: teamsErr } = await supabase.from('teams').select('*');
    if (teamsErr) {
        console.error("Error fetching teams:", teamsErr);
        return;
    }
    const teamNameToId = {};
    for (const t of teams) {
        teamNameToId[t.name?.toLowerCase()] = t.id;
    }
    console.log(`Found ${teams.length} teams: ${teams.map(t => `${t.name} (id=${t.id})`).join(', ')}`);

    // --- Map Excel data ---
    const records = data.map(row => {
        const empCode = cleanCode(row['employee_code'] || row['Employee Code'] || row['Employee_Code']);
        const empName = cleanStr(row['employee_name'] || row['Employee Name'] || row['Employee_Name']);
        const role = cleanStr(row['role'] || row['Role']);
        const supervisorName = cleanStr(row['supervisor_name'] || row['Supervisor Name'] || row['Supervisor_Name']);
        const areaManagerName = cleanStr(row['area_manager_name'] || row['Area Manager Name'] || row['Area_Manager_Name']);
        const teamName = cleanStr(row['team'] || row['Team'] || row['team_name'] || row['Team Name']);
        const teamId = teamName ? (teamNameToId[teamName.toLowerCase()] || null) : null;

        return { employee_code: empCode, employee_name: empName, role, supervisor_name: supervisorName, area_manager_name: areaManagerName, team_id: teamId, _team_name: teamName };
    }).filter(r => r.employee_code);

    // Check unmapped teams
    const unmappedTeams = [...new Set(records.filter(r => r._team_name && !r.team_id).map(r => r._team_name))];
    if (unmappedTeams.length > 0) {
        console.warn(`\n⚠️  WARNING: These team names don't match any team in Supabase:`);
        unmappedTeams.forEach(t => console.warn(`   - "${t}"`));
    }

    // --- Generate SQL ---
    let sql = '-- ===================================================\n';
    sql += '-- Hierarchy replacement SQL\n';
    sql += `-- Generated: ${new Date().toISOString()}\n`;
    sql += `-- Source: hierarchy_export.xlsx (${records.length} records)\n`;
    sql += '-- ===================================================\n\n';
    sql += 'BEGIN;\n\n';
    sql += '-- Step 1: Delete all existing hierarchy rows\n';
    sql += 'DELETE FROM hierarchy;\n\n';
    sql += '-- Step 2: Insert new hierarchy data\n';
    sql += 'INSERT INTO hierarchy (employee_code, employee_name, role, supervisor_name, area_manager_name, team_id)\nVALUES\n';

    const valueRows = records.map(r => {
        return `  (${escSql(r.employee_code)}, ${escSql(r.employee_name)}, ${escSql(r.role)}, ${escSql(r.supervisor_name)}, ${escSql(r.area_manager_name)}, ${r.team_id !== null ? r.team_id : 'NULL'})`;
    });

    sql += valueRows.join(',\n');
    sql += ';\n\n';
    sql += 'COMMIT;\n';

    // Write SQL file
    const outPath = 'E:\\crm extractor\\hierarchy\\hierarchy_replace.sql';
    fs.writeFileSync(outPath, sql, 'utf8');
    console.log(`\n✅ SQL file written to: ${outPath}`);
    console.log(`   ${records.length} INSERT rows generated.`);
    console.log(`\n📋 Copy the contents of that file and paste into the Supabase SQL Editor to run it.`);
}

generateSQL().catch(console.error);
