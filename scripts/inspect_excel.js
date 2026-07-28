const xlsx = require('xlsx');
const wb = xlsx.readFile('E:\\crm extractor\\Periods\\recent\\July 2026 - Focus 1.xlsx');
if (wb.Sheets['RAW DATA']) {
  const data = xlsx.utils.sheet_to_json(wb.Sheets['RAW DATA'], {defval: null});
  console.log('Total RAW DATA Rows:', data.length);
  const userRows = data.filter(r => String(r.user || r.User || '').toLowerCase().includes('abdelrahman arafat'));
  console.log('Abdelrahman Arafat Rows:', userRows.length);
  if (userRows.length > 0) {
    const dates = [...new Set(userRows.map(r => r.date || r.Date))];
    console.log('All unique date strings for Abdelrahman Arafat:', dates);
    const july8Rows = userRows.filter(r => String(r.date || r.Date).includes('2026-07-08') || String(r.date || r.Date).includes('07/08/2026') || String(r.date || r.Date).includes('7/8/2026') || String(r.date || r.Date).includes('08-07-2026'));
    console.log('July 8 rows:', july8Rows.length);
    july8Rows.forEach((r, i) => console.log(i + 1, 'date:', r.date || r.Date, 'shift:', r.shift || r.Shift, 'doctor:', r.doctor_name || r.acc_name));
  }
}
