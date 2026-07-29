/**
 * inspect_excel.js — inspect the RAW DATA sheet of a workbook.
 *
 * FIXED:
 *  - Workbook path, user filter and date filter are arguments instead of
 *    being hardcoded to one file and one person.
 *  - Dates are normalised before comparison, so a single check replaces the
 *    old scattergun of four different string formats — which was itself a
 *    symptom of dates never being normalised on the way into the database.
 *
 * Usage:
 *   node inspect_excel.js "E:\\crm extractor\\Periods\\recent\\July 2026 - Focus 1.xlsx"
 *   node inspect_excel.js "<file>" "abdelrahman arafat" 2026-07-08
 */

const fs = require('fs');
const xlsx = require('xlsx');

const pad = n => String(n).padStart(2, '0');

/** Normalise any date cell to 'YYYY-MM-DD'. */
function isoDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date && !Number.isNaN(v.getTime())) {
        return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    }
    if (typeof v === 'number' && v > 1 && v < 100000) {
        const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
        return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    }
    const s = String(v).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
    if (m) {
        const [, a, b, y] = m;
        const day = Number(a) > 12 ? a : (Number(b) > 12 ? b : a);
        const mon = Number(a) > 12 ? b : (Number(b) > 12 ? a : b);
        return `${y}-${pad(mon)}-${pad(day)}`;
    }
    const d = new Date(s);
    return Number.isNaN(d.getTime())
        ? s
        : `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

const [, , filePath, userFilter, dateFilter] = process.argv;

if (!filePath) {
    console.error('Usage: node inspect_excel.js <workbook.xlsx> [user substring] [YYYY-MM-DD]');
    process.exit(1);
}
if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    process.exit(1);
}

const wb = xlsx.readFile(filePath);
const sheet = wb.Sheets['RAW DATA'] || wb.Sheets['Raw Data'];

if (!sheet) {
    console.error(`No 'RAW DATA' sheet. Sheets present: ${wb.SheetNames.join(', ')}`);
    process.exit(1);
}

const data = xlsx.utils.sheet_to_json(sheet, { defval: null });
console.log(`Total RAW DATA rows: ${data.length}`);

let rows = data;
if (userFilter) {
    const needle = userFilter.toLowerCase();
    rows = data.filter(r => String(r.user || r.User || '').toLowerCase().includes(needle));
    console.log(`Rows matching "${userFilter}": ${rows.length}`);
    if (!rows.length) {
        const names = [...new Set(data.map(r => r.user || r.User).filter(Boolean))].sort();
        console.log(`\nUsers present in this file (${names.length}):`);
        names.forEach(n => console.log(`  ${n}`));
        process.exit(0);
    }
}

const dates = [...new Set(rows.map(r => isoDate(r.date || r.Date)).filter(Boolean))].sort();
console.log(`Unique dates (${dates.length}): ${dates.join(', ')}`);

if (dateFilter) {
    const target = isoDate(dateFilter);
    const dayRows = rows.filter(r => isoDate(r.date || r.Date) === target);
    console.log(`\nRows on ${target}: ${dayRows.length}`);
    dayRows.forEach((r, i) => {
        console.log(
            `  ${i + 1}. shift=${r.shift || r.Shift || '-'}  ` +
            `time=${r.time || r.Time || '-'}  ` +
            `doctor=${r.doctor_name || r.acc_name || '-'}`
        );
    });
}
