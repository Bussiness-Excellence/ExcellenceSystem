const fs = require('fs');
const path = require('path');

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

const headers = {
  "apikey": key,
  "Authorization": `Bearer ${key}`,
  "Content-Type": "application/json"
};

const tables = ["visits", "summaries", "specialty_classification", "product_calls", "coaching_days"];

async function main() {
  for (const t of tables) {
    try {
      const fetch = (await import('node-fetch')).default;
      // Note: we can use standard fetch since Node 18 has global fetch!
      // Let's use global fetch first, fallback to node-fetch if needed.
      const fn = typeof fetch === 'function' ? fetch : global.fetch;
      const res = await fn(`${url}/rest/v1/${t}?select=count&period=eq.July 2026`, { headers });
      if (res.ok) {
        const text = await res.text();
        console.log(`Table ${t} for 'July 2026': ${text}`);
      } else {
        console.log(`Failed to fetch ${t}: ${res.status} ${await res.text()}`);
      }
    } catch (err) {
      console.log(`Error fetching ${t}: ${err.message}`);
    }
  }

  // Check RPC
  try {
    const fn = global.fetch;
    const payload = {
      "p_period": "July 2026",
      "p_codes": ["7535"],
      "p_is_admin": true,
      "p_is_manager": true,
      "p_start_date": "2026-07-01",
      "p_end_date": "2026-07-31"
    };
    const res = await fn(`${url}/rest/v1/rpc/get_dashboard_data`, {
      method: 'POST',
      headers,
      body: JSON.stringify(payload)
    });
    if (res.ok) {
      const data = await res.json();
      console.log("\nRPC get_dashboard_data keys returned:", Object.keys(data));
      for (const [k, v] of Object.entries(data)) {
        console.log(`  ${k}: ${Array.isArray(v) ? v.length : typeof v}`);
      }
    } else {
      console.log(`\nFailed to call RPC get_dashboard_data: ${res.status} ${await res.text()}`);
    }
  } catch (err) {
    console.log(`Error calling RPC: ${err.message}`);
  }
}

main();
