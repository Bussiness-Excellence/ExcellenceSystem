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

async function updateHierarchy() {
    // ===== UPDATED FILE PATH =====
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

    // --- Step 1: Fetch the teams table to map team names -> team IDs ---
    console.log("\nFetching teams table...");
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

    // --- Step 2: Map Excel data to DB format ---
    const excelRecords = data.map(row => {
        const empCode = cleanCode(row['employee_code'] || row['Employee Code'] || row['Employee_Code']);
        const empName = cleanStr(row['employee_name'] || row['Employee Name'] || row['Employee_Name']);
        const role = cleanStr(row['role'] || row['Role']);
        const supervisorName = cleanStr(row['supervisor_name'] || row['Supervisor Name'] || row['Supervisor_Name']);
        const areaManagerName = cleanStr(row['area_manager_name'] || row['Area Manager Name'] || row['Area_Manager_Name']);
        const teamName = cleanStr(row['team'] || row['Team'] || row['team_name'] || row['Team Name']);
        const division = cleanStr(row['division'] || row['Division']);

        // Resolve team_id from team name
        const teamId = teamName ? (teamNameToId[teamName.toLowerCase()] || null) : null;

        return {
            employee_code: empCode,
            employee_name: empName,
            role: role,
            supervisor_name: supervisorName,
            area_manager_name: areaManagerName,
            team_id: teamId,
            _team_name: teamName, // keep for logging (won't be sent to DB)
            _division: division,  // keep for logging
        };
    });

    // Check for team names that couldn't be resolved
    const unmappedTeams = [...new Set(
        excelRecords
            .filter(r => r._team_name && !r.team_id)
            .map(r => r._team_name)
    )];
    if (unmappedTeams.length > 0) {
        console.warn(`\n⚠️  WARNING: These team names from Excel don't match any team in Supabase:`);
        unmappedTeams.forEach(t => console.warn(`   - "${t}"`));
        console.warn(`   They will be inserted with team_id = null.`);
    }

    // --- Step 3: Delete ALL existing hierarchy rows ---
    console.log("\n🗑️  Deleting ALL existing hierarchy rows...");
    const { error: deleteErr, count: deleteCount } = await supabase
        .from('hierarchy')
        .delete({ count: 'exact' })
        .not('id', 'is', null);  // matches all rows

    if (deleteErr) {
        console.error("Error deleting existing hierarchy:", deleteErr.message);
        console.log("Trying alternative delete method...");
        
        // Alternative: delete in batches by fetching IDs first
        const { data: allRows, error: fetchErr } = await supabase
            .from('hierarchy')
            .select('id')
            .range(0, 10000);
        
        if (fetchErr) {
            console.error("Error fetching hierarchy IDs:", fetchErr.message);
            return;
        }
        
        if (allRows && allRows.length > 0) {
            console.log(`Found ${allRows.length} rows to delete...`);
            for (let i = 0; i < allRows.length; i += 500) {
                const batch = allRows.slice(i, i + 500).map(r => r.id);
                const { error: batchDelErr } = await supabase
                    .from('hierarchy')
                    .delete()
                    .in('id', batch);
                if (batchDelErr) {
                    console.error(`Error deleting batch ${i}:`, batchDelErr.message);
                } else {
                    console.log(`  ✅ Deleted batch ${i} to ${i + batch.length}`);
                }
            }
        }
    } else {
        console.log(`  ✅ Deleted ${deleteCount || 'all'} existing hierarchy rows.`);
    }

    // --- Step 4: Insert all new records ---
    const cleanRecords = excelRecords
        .filter(r => r.employee_code) // skip rows without employee code
        .map(r => ({
            employee_code: r.employee_code,
            employee_name: r.employee_name,
            role: r.role,
            supervisor_name: r.supervisor_name,
            area_manager_name: r.area_manager_name,
            team_id: r.team_id,
        }));

    console.log(`\n📤 Inserting ${cleanRecords.length} new records...`);
    let totalInserted = 0;
    for (let i = 0; i < cleanRecords.length; i += 500) {
        const chunk = cleanRecords.slice(i, i + 500);
        const { error } = await supabase.from('hierarchy').insert(chunk);
        if (error) {
            console.error(`  ❌ Error inserting batch ${i} to ${i + chunk.length}:`, error.message);
            // Log the first record of the failing batch for debugging
            console.error(`  First record in batch:`, JSON.stringify(chunk[0], null, 2));
        } else {
            totalInserted += chunk.length;
            console.log(`  ✅ Inserted batch ${i} to ${i + chunk.length}`);
        }
    }

    console.log(`\n🎉 Done! Inserted ${totalInserted} / ${cleanRecords.length} records into hierarchy table.`);
}

updateHierarchy().catch(console.error);
