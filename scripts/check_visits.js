/**
 * check_visits.js — quick lookup of a user's visits on a given date.
 *
 * FIXED:
 *  - User and date are arguments rather than being hardcoded to two people
 *    and one day.
 *  - Uses the shared config loader, so credentials and the missing-env error
 *    message are consistent with the rest of the scripts.
 *  - Reports query errors instead of silently printing "undefined".
 *
 * Usage:
 *   node check_visits.js "Maher Abd Elfattah" 2026-07-01
 *   node check_visits.js "Maher Abd Elfattah" 2026-07-01 "Abdullah ahmed mohamed Abdelaal"
 */

const { getAnonClient } = require('./config');

const supabase = getAnonClient();

const [, , ...args] = process.argv;
const date = args.find(a => /^\d{4}-\d{2}-\d{2}$/.test(a));
const users = args.filter(a => a !== date);

if (!users.length || !date) {
    console.error('Usage: node check_visits.js "<user name>" <YYYY-MM-DD> ["<another user>"]');
    process.exit(1);
}

async function check() {
    for (const user of users) {
        const { data, error } = await supabase
            .from('visits')
            .select('*')
            .eq('user', user)
            .eq('visit_date', date);

        if (error) {
            console.error(`Query failed for ${user}: ${error.message}`);
            continue;
        }

        console.log(`\n${user} — ${date}: ${data.length} visit(s)`);
        data.forEach((v, i) => {
            console.log(
                `  ${i + 1}. ${v.visit_time || '-'}  ${v.shift || '-'}  ` +
                `${v.doctor_name || v.acc_name || '-'}  [batch: ${v.upload_batch || '-'}]`
            );
        });

        // Repeated identical rows here are the signature of the duplicate bug
        // fixed in pulpoplus_upload_to_supabase.py and emergency_upload.js.
        const seen = new Map();
        for (const v of data) {
            const k = `${v.visit_time}|${v.acc_id}|${v.doctor_key}`;
            seen.set(k, (seen.get(k) || 0) + 1);
        }
        const dupes = [...seen.entries()].filter(([, n]) => n > 1);
        if (dupes.length) {
            console.warn(`  WARNING: ${dupes.length} duplicated visit(s) for this user/date.`);
        }
    }
}

check().catch(err => {
    console.error('Unexpected failure:', err);
    process.exitCode = 1;
});
