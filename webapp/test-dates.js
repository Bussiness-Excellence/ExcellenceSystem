#!/usr/bin/env node
/**
 * Smoke test for the Time Slicer date logic.
 *
 * Rather than re-implementing normalizeDateStr (which would drift from the real
 * thing), this reads src/components/Dashboard.js, extracts the function body and
 * the GRAIN_RANGES table, and runs the actual shipped source.
 *
 *   node test-dates.js [path/to/Dashboard.js]
 *
 * Exits non-zero on failure, so it works as a pre-push hook or CI step.
 */
const fs = require('fs');
const path = process.argv[2] || 'src/components/Dashboard.js';

if (!fs.existsSync(path)) {
  console.error(`Cannot find ${path} — pass the path as an argument.`);
  process.exit(2);
}
const src = fs.readFileSync(path, 'utf8');

// ---- extract GRAIN_RANGES -------------------------------------------------
const rangesMatch = src.match(/const GRAIN_RANGES = React\.useMemo\(\(\) => \(([\s\S]*?)\), \[\]\);/);
if (!rangesMatch) {
  console.error('FAIL: GRAIN_RANGES not found. Did the patch get applied?');
  process.exit(1);
}
const GRAIN_RANGES = eval(`(${rangesMatch[1]})`);

// ---- extract normalizeDateStr --------------------------------------------
const normMatch = src.match(/const normalizeDateStr = useCallback\((\(dStr\) => \{[\s\S]*?\n  \})\, \[\]\);/);
if (!normMatch) {
  console.error('FAIL: normalizeDateStr not found or its shape changed.');
  process.exit(1);
}
const normalizeDateStr = eval(`(${normMatch[1]})`);

// ---- guard: the old estimator must be gone --------------------------------
let failures = 0;
const check = (label, cond, detail = '') => {
  if (cond) { console.log(`  ok   ${label}`); }
  else { console.log(`  FAIL ${label}${detail ? ' — ' + detail : ''}`); failures++; }
};

console.log('\nRegression guards');
check('getTimeGrainRatio removed', !/const getTimeGrainRatio\s*=/.test(src));
check('ratio-scaling removed', !/Math\.max\(1, Math\.round\(x\[sk\] \* ratio\)\)/.test(src));
check('filter fails closed', /if \(!isoDate\) return false;/.test(src));
check('specialty fallback gated', /Same as specialty: the product aggregate has no date column/.test(src));

console.log('\nnormalizeDateStr');
const cases = [
  ['2026-07-06', '2026-07-06', 'ISO passthrough'],
  ['2026-7-6',   '2026-07-06', 'ISO unpadded'],
  ['2026/07/06', '2026-07-06', 'ISO with slashes'],
  ['2026-02-30', '',           'invalid day rejected'],
  ['2026-13-01', '',           'invalid month rejected'],
  ['2026-02-29', '',           '2026 is not a leap year'],
  ['2024-02-29', '2024-02-29', '2024 leap day valid'],
  ['25/07/2026', '2026-07-25', 'DD/MM when unambiguous'],
  ['07/25/2026', '2026-07-25', 'MM/DD when unambiguous'],
  ['07/01/2026', '',           'ambiguous -> refuse to guess'],
  ['',           '',           'empty'],
  [null,         '',           'null'],
  ['not a date', '',           'garbage'],
];
for (const [input, expected, label] of cases) {
  const got = normalizeDateStr(input);
  check(label, got === expected, `${JSON.stringify(input)} -> ${JSON.stringify(got)}, wanted ${JSON.stringify(expected)}`);
}

console.log('\nGRAIN_RANGES partition');
const days = Array.from({ length: 31 }, (_, i) => i + 1);
const inRange = (d, g) => d >= GRAIN_RANGES[g][0] && d <= GRAIN_RANGES[g][1];

for (const group of [['week1', 'week2', 'week3', 'week4'], ['biweekly1', 'biweekly2']]) {
  const counts = days.map(d => group.filter(g => inRange(d, g)).length);
  check(`${group.join('+')} covers every day exactly once`,
    counts.every(c => c === 1),
    `days covered ${counts.filter(c => c === 1).length}/31, overlaps on ${days.filter((d, i) => counts[i] > 1)}`);
}

// week4 must absorb the tail of long months
check('week4 reaches day 31', GRAIN_RANGES.week4[1] === 31);
check('biweekly2 reaches day 31', GRAIN_RANGES.biweekly2[1] === 31);

console.log(failures === 0 ? '\nAll checks passed.\n' : `\n${failures} check(s) failed.\n`);
process.exit(failures === 0 ? 0 : 1);
