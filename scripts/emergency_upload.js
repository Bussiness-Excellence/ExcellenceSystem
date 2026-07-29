/**
 * emergency_upload.js — bulk restore of dashboard data from Excel workbooks.
 *
 * FIXED:
 *  - Service-role key no longer hardcoded; read from .env via config.js.
 *  - THE DUPLICATE BUG: this script cleared each batch with
 *        .eq('upload_batch', 'recent')
 *    but pulpoplus_auto_upload.py writes rows tagged
 *        recent_July 2026 - Focus 1
 *    (batch + "_" + filename). Those never match, so the clear step deleted
 *    nothing and every re-run stacked another full copy of the data on top.
 *    Clearing is now done with a LIKE prefix match that covers both shapes.
 *  - Dates are normalised to YYYY-MM-DD instead of being String()'d straight
 *    from a JS Date, which produced "Wed Jul 01 2026 00:00:00 GMT+0200 ...".
 *  - Excel "~$" lock files are skipped.
 *  - Insert failures now stop the run instead of scrolling past.
 *  - Folder paths come from .env.
 */

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { getServiceClient, envPath } = require('./config');

const supabase = getServiceClient();

// ── helpers ────────────────────────────────────────────────────────────────

const cleanStr = v =>
    (v === null || v === undefined || v === '' || v === 'NaN') ? null : String(v).trim();

const cleanNum = v =>
    (v === null || v === undefined || v === '' || Number.isNaN(Number(v))) ? null : Number(v);

const cleanCode = v => {
    let s = cleanStr(v);
    if (!s) return null;
    if (s.endsWith('.0')) s = s.slice(0, -2);
    return s;
};

const pad = n => String(n).padStart(2, '0');

/**
 * Normalise any date cell to an ISO 'YYYY-MM-DD' string.
 * sheet_to_json returns a real Date for date-formatted cells, and String()
 * on that gives a long locale string Postgres rejects or misparses.
 */
const cleanDate = v => {
    if (v === null || v === undefined || v === '') return null;

    if (v instanceof Date && !Number.isNaN(v.getTime())) {
        return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    }

    // Excel serial number (days since 1899-12-30)
    if (typeof v === 'number' && v > 1 && v < 100000) {
        const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
        return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    }

    const s = String(v).trim();
    if (!s || s.toLowerCase() === 'nan') return null;

    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);       // already ISO

    let m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);          // dd/mm/yyyy
    if (m) {
        let [, a, b, y] = m;
        // Disambiguate: a value above 12 must be the day.
        const day = Number(a) > 12 ? a : (Number(b) > 12 ? b : a);
        const mon = Number(a) > 12 ? b : (Number(b) > 12 ? a : b);
        return `${y}-${pad(mon)}-${pad(day)}`;
    }

    const parsed = new Date(s);
    if (!Number.isNaN(parsed.getTime())) {
        return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
    }
    return s;
};

const cleanTime = v => {
    if (v instanceof Date && !Number.isNaN(v.getTime())) {
        return `${pad(v.getHours())}:${pad(v.getMinutes())}:${pad(v.getSeconds())}`;
    }
    const s = cleanStr(v);
    if (!s) return null;
    // '2026-07-01 09:30:00' -> '09:30:00'
    const m = s.match(/\b(\d{1,2}:\d{2}(?::\d{2})?)\b/);
    return m ? m[1] : s;
};

const hmToMin = v => {
    const s = cleanStr(v);
    if (!s || !s.includes(':')) return null;
    const [h, m] = s.split(':');
    const hh = parseInt(h, 10);
    const mm = parseInt(m, 10);
    if (Number.isNaN(hh) || Number.isNaN(mm)) return null;
    return hh * 60 + mm;
};

let hadError = false;

/** Insert in chunks; report and flag failures rather than silently continuing. */
async function insertChunks(table, rows) {
    if (!rows.length) return 0;
    let inserted = 0;
    for (let i = 0; i < rows.length; i += 500) {
        const chunk = rows.slice(i, i + 500);
        const { error } = await supabase.from(table).insert(chunk);
        if (error) {
            hadError = true;
            console.error(`  ERROR inserting into ${table} (rows ${i}-${i + chunk.length}): ${error.message}`);
            console.error(`  First failing record:`, JSON.stringify(chunk[0], null, 2));
            return inserted;
        }
        inserted += chunk.length;
    }
    console.log(`  ${table}: ${inserted} rows`);
    return inserted;
}

// ── sheet handlers ─────────────────────────────────────────────────────────

const SHEET_MAP = [
    {
        sheet: 'Summary',
        table: 'summaries',
        map: (r, period, batch) => ({
            employee_code: cleanCode(r['Employee Code']),
            user_name: cleanStr(r['User']),
            period,
            team: cleanStr(r['Team']),
            territory: cleanStr(r['Territory']),
            is_manager: String(r['Is Manager'] || '').trim().toLowerCase() === 'yes',
            working_days: cleanNum(r['Working Days']),
            complete_field_days: cleanNum(r['Complete Field Days']),
            office_work_days: cleanNum(r['Office Work Days']),
            no_activities: cleanNum(r['No. of Activities']),
            no_events: cleanNum(r['No. of Events']),
            double_visit_days: cleanNum(r['Double Visit Days']),
            coaching_days: cleanNum(r['Coaching Days']),
            am_shift_days: cleanNum(r['AM Shift Days']),
            pm_shift_days: cleanNum(r['PM Shift Days']),
            avg_am_duration_min: hmToMin(r['Avg AM Shift Duration (h:mm)']),
            avg_pm_duration_min: hmToMin(r['Avg PM Shift Duration (h:mm)']),
            upload_batch: batch,
        }),
    },
    {
        sheet: 'Specialty x Class',
        table: 'specialty_classification',
        map: (r, period, batch) => ({
            employee_code: cleanCode(r['Employee Code']),
            user_name: cleanStr(r['User']),
            period,
            specialty: cleanStr(r['Specialty']),
            classification: cleanStr(r['Classification']),
            shift: cleanStr(r['Shift']),
            call_count: cleanNum(r['Call Count']),
            unique_doctors: cleanNum(r['Unique Doctors']),
            upload_batch: batch,
        }),
    },
    {
        sheet: 'Coaching Days',
        table: 'coaching_days',
        map: (r, period, batch) => ({
            manager_name: cleanStr(r['Manager']),
            manager_code: cleanCode(r['Manager Code']),
            rep_name: cleanStr(r['Rep']),
            rep_code: cleanCode(r['Rep Code']),
            coaching_date: cleanDate(r['Date']),
            team: cleanStr(r['Team']),
            am_visits: cleanNum(r['AM Visits']),
            am_accompanied: cleanNum(r['AM Accompanied']),
            pm_visits: cleanNum(r['PM Visits']),
            pm_accompanied: cleanNum(r['PM Accompanied']),
            upload_batch: batch,
        }),
    },
    {
        sheet: 'Product Calls per spec',
        table: 'product_calls',
        map: (r, period, batch) => ({
            employee_code: cleanCode(r['Employee Code']),
            user_name: cleanStr(r['User']),
            period,
            specialty: cleanStr(r['Specialty']),
            product: cleanStr(r['Product']),
            shift: cleanStr(r['Shift']),
            call_count: cleanNum(r['Call Count']),
            unique_doctors: cleanNum(r['Unique Doctors']),
            upload_batch: batch,
        }),
    },
];

function mapVisits(r, period, batch) {
    return {
        team: cleanStr(r['Team']),
        user: cleanStr(r['user'] || r['User']),
        employee_code: cleanCode(r['user_code'] || r['Employee Code']),
        territory: cleanStr(r['territory'] || r['Territory']),
        visit_date: cleanDate(r['date'] || r['Date']),
        visit_time: cleanTime(r['time'] || r['Time']),
        acc_type_raw: cleanStr(r['acc_type_raw']),
        acc_type_category: cleanStr(r['acc_type_category']),
        shift: cleanStr(r['shift'] || r['Shift']),
        visit_type_raw: cleanStr(r['visit_type_raw']),
        visit_type_category: cleanStr(r['visit_type_category']),
        acc_id: cleanStr(r['acc_id']),
        acc_name: cleanStr(r['acc_name']),
        doctor_key: cleanStr(r['doctor_key']),
        doctor_name: cleanStr(r['doctor_name']),
        specialty: cleanStr(r['specialty'] || r['Specialty']),
        classification: cleanStr(r['classification'] || r['Classification']),
        products: cleanStr(r['products'] || r['Products']),
        notes: cleanStr(r['notes']),
        upload_batch: batch,
    };
}

async function uploadFile(filePath, periodLabel, batchLabel) {
    console.log(`\nReading ${path.basename(filePath)}...`);
    let wb;
    try {
        wb = xlsx.readFile(filePath);
    } catch (err) {
        hadError = true;
        console.error(`  Could not read workbook: ${err.message}`);
        return;
    }

    for (const { sheet, table, map } of SHEET_MAP) {
        if (!wb.Sheets[sheet]) continue;
        const data = xlsx.utils.sheet_to_json(wb.Sheets[sheet], { defval: null });
        if (!data.length) continue;
        await insertChunks(table, data.map(r => map(r, periodLabel, batchLabel)));
    }

    const rawSheet = wb.Sheets['RAW DATA'] || wb.Sheets['Raw Data'];
    if (rawSheet) {
        const data = xlsx.utils.sheet_to_json(rawSheet, { defval: null });
        if (data.length) {
            await insertChunks('visits', data.map(r => mapVisits(r, periodLabel, batchLabel)));
        }
    }
}

const TABLES = ['summaries', 'specialty_classification', 'product_calls', 'coaching_days', 'visits'];

/**
 * Clear a batch.
 *
 * Rows written by this script are tagged exactly `batchLabel`, but rows
 * written by pulpoplus_auto_upload.py are tagged `batchLabel_<filename>`.
 * An .eq() match only caught the first kind — which is why re-running left
 * the old rows in place and doubled the dashboard numbers.
 */
async function clearBatch(batchLabel) {
    console.log(`\n=== Clearing batch '${batchLabel}' (and '${batchLabel}_*') ===`);
    for (const table of TABLES) {
        const { error } = await supabase
            .from(table)
            .delete()
            .or(`upload_batch.eq.${batchLabel},upload_batch.like.${batchLabel}\\_%`);
        if (error) {
            hadError = true;
            console.error(`  Could not clear ${table}: ${error.message}`);
        }
    }
}

async function runBatch(folder, periodLabel, batchLabel) {
    if (!fs.existsSync(folder)) {
        console.log(`\nFolder not found, skipping: ${folder}`);
        return;
    }

    const files = fs.readdirSync(folder)
        .filter(f => f.toLowerCase().endsWith('.xlsx') && !f.startsWith('~$') && !f.startsWith('.'));

    if (!files.length) {
        console.log(`\nNo workbooks in ${folder}, skipping.`);
        return;
    }

    await clearBatch(batchLabel);
    for (const f of files) {
        await uploadFile(path.join(folder, f), periodLabel, batchLabel);
    }
}

async function main() {
    const root = envPath('CRM_ROOT', 'E:\\crm extractor');

    const batches = [
        [path.join(root, 'june'), 'June 2026', 'june_2026'],
        [path.join(root, 'july'), 'July 2026', 'july_2026'],
        [envPath('PERIOD_LAST_MONTH_DIR', path.join(root, 'Periods', 'last_month')), 'Last Month', 'last_month'],
        [envPath('PERIOD_RECENT_DIR', path.join(root, 'Periods', 'recent')), 'Recent', 'recent'],
    ];

    for (const [folder, period, batch] of batches) {
        await runBatch(folder, period, batch);
    }

    if (hadError) {
        console.error('\nFinished WITH ERRORS — the data above is incomplete. Review the messages.');
        process.exitCode = 1;
    } else {
        console.log('\nDone restoring dashboard data.');
    }
}

main().catch(err => {
    console.error('Unexpected failure:', err);
    process.exitCode = 1;
});
