const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { createClient } = require('@supabase/supabase-js');

const supabaseUrl = 'https://xxbfwvlqixnmonxytdxq.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh4YmZ3dmxxaXhubW9ueHl0ZHhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4Mjc1NjE2NSwiZXhwIjoyMDk4MzMyMTY1fQ.PSk6RyFmg_OFTcCtYO74AeJj6wT4FGZS2K2JT9GEJ_A';
const supabase = createClient(supabaseUrl, supabaseKey);

// Helpers
const cleanStr = v => (v === null || v === undefined || v === '' || v === 'NaN') ? null : String(v).trim();
const cleanNum = v => (v === null || v === undefined || v === '' || Number.isNaN(Number(v))) ? null : Number(v);
const cleanCode = v => {
    let s = cleanStr(v);
    if (!s) return null;
    if (s.endsWith('.0')) s = s.slice(0, -2);
    return s;
};
const hmToMin = v => {
    const s = cleanStr(v);
    if (!s || !s.includes(':')) return null;
    const [h, m] = s.split(':');
    return parseInt(h) * 60 + parseInt(m);
};

async function uploadFile(filePath, periodLabel, batchLabel) {
    console.log(`\nReading ${filePath}...`);
    const wb = xlsx.readFile(filePath);
    
    // Summary
    if (wb.Sheets['Summary']) {
        const data = xlsx.utils.sheet_to_json(wb.Sheets['Summary'], { defval: null });
        if (data.length > 0) {
            console.log(`Uploading ${data.length} summaries...`);
            const rows = data.map(r => ({
                employee_code: cleanCode(r['Employee Code']),
                user_name: cleanStr(r['User']),
                period: periodLabel,
                team: cleanStr(r['Team']),
                territory: cleanStr(r['Territory']),
                is_manager: String(r['Is Manager']||'').trim().toLowerCase() === 'yes',
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
                upload_batch: batchLabel,
            }));
            
            for (let i = 0; i < rows.length; i += 500) {
                const chunk = rows.slice(i, i + 500);
                const { error } = await supabase.from('summaries').insert(chunk);
                if (error) console.error("Error inserting summaries:", error);
            }
        }
    }

    // Specialty x Class
    if (wb.Sheets['Specialty x Class']) {
        const data = xlsx.utils.sheet_to_json(wb.Sheets['Specialty x Class'], { defval: null });
        if (data.length > 0) {
            console.log(`Uploading ${data.length} specialties...`);
            const rows = data.map(r => ({
                employee_code: cleanCode(r['Employee Code']),
                user_name: cleanStr(r['User']),
                period: periodLabel,
                specialty: cleanStr(r['Specialty']),
                classification: cleanStr(r['Classification']),
                shift: cleanStr(r['Shift']),
                call_count: cleanNum(r['Call Count']),
                unique_doctors: cleanNum(r['Unique Doctors']),
                upload_batch: batchLabel,
            }));
            
            for (let i = 0; i < rows.length; i += 500) {
                const chunk = rows.slice(i, i + 500);
                const { error } = await supabase.from('specialty_classification').insert(chunk);
                if (error) console.error("Error inserting specialties:", error);
            }
        }
    }

    // Coaching Days
    if (wb.Sheets['Coaching Days']) {
        const data = xlsx.utils.sheet_to_json(wb.Sheets['Coaching Days'], { defval: null });
        if (data.length > 0) {
            console.log(`Uploading ${data.length} coaching days...`);
            const rows = data.map(r => ({
                manager_name: cleanStr(r['Manager']),
                manager_code: cleanCode(r['Manager Code']),
                rep_name: cleanStr(r['Rep']),
                rep_code: cleanCode(r['Rep Code']),
                coaching_date: cleanStr(r['Date']),
                team: cleanStr(r['Team']),
                am_visits: cleanNum(r['AM Visits']),
                am_accompanied: cleanNum(r['AM Accompanied']),
                pm_visits: cleanNum(r['PM Visits']),
                pm_accompanied: cleanNum(r['PM Accompanied']),
                upload_batch: batchLabel,
            }));
            
            for (let i = 0; i < rows.length; i += 500) {
                const chunk = rows.slice(i, i + 500);
                const { error } = await supabase.from('coaching_days').insert(chunk);
                if (error) console.error("Error inserting coaching:", error);
            }
        }
    }

    // Product Calls per spec
    if (wb.Sheets['Product Calls per spec']) {
        const data = xlsx.utils.sheet_to_json(wb.Sheets['Product Calls per spec'], { defval: null });
        if (data.length > 0) {
            console.log(`Uploading ${data.length} products...`);
            const rows = data.map(r => ({
                employee_code: cleanCode(r['Employee Code']),
                user_name: cleanStr(r['User']),
                period: periodLabel,
                specialty: cleanStr(r['Specialty']),
                product: cleanStr(r['Product']),
                shift: cleanStr(r['Shift']),
                call_count: cleanNum(r['Call Count']),
                unique_doctors: cleanNum(r['Unique Doctors']),
                upload_batch: batchLabel,
            }));
            
            for (let i = 0; i < rows.length; i += 500) {
                const chunk = rows.slice(i, i + 500);
                const { error } = await supabase.from('product_calls').insert(chunk);
                if (error) console.error("Error inserting products:", error);
            }
        }
    }

    // RAW DATA -> visits
    const rawSheet = wb.Sheets['RAW DATA'] || wb.Sheets['Raw Data'];
    if (rawSheet) {
        const data = xlsx.utils.sheet_to_json(rawSheet, { defval: null });
        if (data.length > 0) {
            console.log(`Uploading ${data.length} raw visits...`);
            const rows = data.map(r => ({
                team: cleanStr(r['Team']),
                user: cleanStr(r['user'] || r['User']),
                employee_code: cleanCode(r['user_code'] || r['Employee Code']),
                territory: cleanStr(r['territory'] || r['Territory']),
                visit_date: cleanStr(r['date'] || r['Date']),
                visit_time: cleanStr(r['time'] || r['Time']),
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
                upload_batch: batchLabel,
            }));
            
            for (let i = 0; i < rows.length; i += 500) {
                const chunk = rows.slice(i, i + 500);
                const { error } = await supabase.from('visits').insert(chunk);
                if (error) console.error("Error inserting visits:", error);
            }
        }
    }
}

async function runBatch(folder, periodLabel, batchLabel) {
    console.log(`\n=== Clearing batch '${batchLabel}' ===`);
    // Clear old data
    await supabase.from('summaries').delete().eq('upload_batch', batchLabel);
    await supabase.from('specialty_classification').delete().eq('upload_batch', batchLabel);
    await supabase.from('product_calls').delete().eq('upload_batch', batchLabel);
    await supabase.from('coaching_days').delete().eq('upload_batch', batchLabel);
    await supabase.from('visits').delete().eq('upload_batch', batchLabel);

    if (!fs.existsSync(folder)) return;
    const files = fs.readdirSync(folder).filter(f => f.endsWith('.xlsx') && !f.startsWith('~'));
    for (const f of files) {
        await uploadFile(path.join(folder, f), periodLabel, batchLabel);
    }
}

async function main() {
    await runBatch('e:\\crm extractor\\june', 'June 2026', 'june_2026');
    await runBatch('e:\\crm extractor\\july', 'July 2026', 'july_2026');
    await runBatch('e:\\crm extractor\\Periods\\last_month', 'Last Month', 'last_month');
    await runBatch('e:\\crm extractor\\Periods\\recent', 'Recent', 'recent');
    console.log("\nDone restoring dashboard data!");
}

main().catch(console.error);
