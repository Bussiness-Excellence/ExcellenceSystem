const fs = require('fs');
const path = require('path');
const { createClient } = require('@supabase/supabase-js');

// Read dotenv
const envContent = fs.readFileSync(path.join(__dirname, '.env'), 'utf8');
const env = {};
envContent.split('\n').forEach(line => {
  const parts = line.split('=');
  if (parts.length >= 2) {
    const k = parts[0].trim();
    const v = parts.slice(1).join('=').trim().replace(/^['"]|['"]$/g, '');
    env[k] = v;
  }
});

const url = env.SUPABASE_URL;
const key = env.SUPABASE_SERVICE_ROLE_KEY;

const supabase = createClient(url, key);

async function main() {
  console.log("Fetching summaries sample...");
  const { data: summaries, error: sErr } = await supabase.from('summaries').select('*').limit(5);
  if (sErr) console.log("Error summaries:", sErr);
  else console.log(`Fetched ${summaries.length} summaries. Sample:`, summaries);

  console.log("\nFetching hierarchy sample...");
  const { data: hierarchy, error: hErr } = await supabase.from('hierarchy').select('*').limit(5);
  if (hErr) console.log("Error hierarchy:", hErr);
  else console.log(`Fetched ${hierarchy.length} hierarchy records. Sample:`, hierarchy);

  const codes = (hierarchy || []).map(h => h.employee_code).filter(Boolean);
  console.log(`\nCodes count from hierarchy: ${codes.length}`);

  console.log("\nCalling get_dashboard_data RPC...");
  const { data, error } = await supabase.rpc('get_dashboard_data', {
    p_period: "July 2026",
    p_codes: codes,
    p_is_admin: true,
    p_is_manager: true,
    p_start_date: "2026-07-01",
    p_end_date: "2026-07-31"
  });

  if (error) {
    console.log("RPC Error:", error);
  } else {
    console.log("RPC Success. Returned keys:", Object.keys(data));
    for (const [k, v] of Object.entries(data)) {
      console.log(`  ${k}: ${Array.isArray(v) ? v.length : typeof v}`);
    }
    if (data.summaries && data.summaries.length > 0) {
      console.log("Sample summary record from RPC:", data.summaries[0]);
    }
  }
}

main();
