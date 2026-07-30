/**
 * check_visits.js — quick lookup of a user's visits on a given date.
 *
 * Uses native Node fetch & config helpers (zero npm dependencies required).
 *
 * Usage:
 *   node scripts/check_visits.js "Maher Abd Elfattah" 2026-07-01
 *   node scripts/check_visits.js "Maher Abd Elfattah" 2026-07-01 "Abdullah ahmed mohamed Abdelaal"
 */

const { getUrl, getAnonKey } = require('./config');

const supabaseUrl = getUrl();
const supabaseKey = getAnonKey();

const [, , ...args] = process.argv;
const date = args.find(a => /^\d{4}-\d{2}-\d{2}$/.test(a));
const users = args.filter(a => a !== date);

if (!users.length || !date) {
    console.error('Usage: node check_visits.js "<user name>" <YYYY-MM-DD> ["<another user>"]');
    process.exit(1);
}

async function check() {
    for (const user of users) {
        const queryUrl = `${supabaseUrl}/rest/v1/visits?user=eq.${encodeURIComponent(user)}&visit_date=eq.${encodeURIComponent(date)}&select=*`;
        
        const res = await fetch(queryUrl, {
            headers: {
                'apikey': supabaseKey,
                'Authorization': `Bearer ${supabaseKey}`
            }
        });

        if (!res.ok) {
            const errText = await res.text();
            console.error(`Query failed for ${user}: ${res.status} ${errText}`);
            continue;
        }

        const data = await res.json();

        console.log(`\n${user} — ${date}: ${data.length} visit(s)`);
        data.forEach((v, i) => {
            console.log(
                `  ${i + 1}. ${v.visit_time || '-'}  ${v.shift || '-'}  ` +
                `${v.doctor_name || v.acc_name || '-'}  [batch: ${v.upload_batch || '-'}]`
            );
        });

        // Repeated identical rows here are the signature of the duplicate bug
        const seen = new Map();
        for (const v of data) {
            const k = `${v.visit_time}|${v.acc_id}|${v.doctor_key}`;
            seen.set(k, (seen.get(k) || 0) + 1);
        }
        const dupes = [...seen.entries()].filter(([, n]) => n > 1);
        if (dupes.length) {
            console.warn(`  ⚠️ WARNING: ${dupes.length} duplicated visit(s) for this user/date.`);
        } else if (data.length > 0) {
            console.log(`  ✅ No duplicate visits found.`);
        }
    }
}

check().catch(err => {
    console.error('Unexpected failure:', err);
    process.exitCode = 1;
});
