/**
 * salient_theme.js — apply the Salient premium navy dark theme.
 *
 * FIXED: every substitution now reports whether it matched, the files are
 * backed up before writing, and the script is idempotent (safe to re-run).
 * The old version printed a success message unconditionally.
 */

const { CssEditor } = require('./css_utils');

console.log('Applying Salient Premium Navy theme...');

// ── src/index.css ─────────────────────────────────────────────────────────
const index = new CssEditor('src/index.css');

index.sub(
    'body background -> navy radial gradient',
    /background:\s*#0B1120;/,
    'background: radial-gradient(ellipse at center, #101931 0%, #050811 100%); background-attachment: fixed;'
);

index.sub(
    'overlay gradient -> soft navy bloom',
    /linear-gradient\(160deg, transparent 0%, rgba\(255,255,255,0\.01\) 100%\)/,
    'radial-gradient(circle at 50% 30%, rgba(20, 35, 70, 0.4) 0%, transparent 60%)'
);

// ── src/pages/Dashboard.css ───────────────────────────────────────────────
const dash = new CssEditor('src/pages/Dashboard.css');

dash.sub(
    'html.dark variables -> navy palette',
    /html\.dark\s*\{[\s\S]*?--sh3:[^\n]+\n\}/,
    `html.dark {
  --navy:   #f8fafc;
  --navy2:  #e2e8f0;
  --navy3:  #cbd5e1;
  --navy4:  #94a3b8;
  --gold:   #ffffff;
  --gold2:  #e2e8f0;
  --gold-glow: rgba(255, 255, 255, 0.2);
  --bg:     transparent; /* Let the body radial gradient show through */
  --surf:   rgba(12, 20, 38, 0.65); /* Sleek translucent navy panels */
  --bdr:    rgba(255, 255, 255, 0.08);
  --bdr2:   rgba(255, 255, 255, 0.12);
  --txt:    #ffffff;
  --mute:   rgba(255, 255, 255, 0.6);
  --am:     #38bdf8;
  --am-bg:  rgba(56, 189, 248, 0.15);
  --pm:     #34d399;
  --pm-bg:  rgba(52, 211, 153, 0.15);
  --sh: 0 4px 16px rgba(0, 0, 0, 0.5);
  --sh2: 0 8px 24px rgba(0, 0, 0, 0.6);
  --sh3: 0 12px 36px rgba(0, 0, 0, 0.8);
}`
);

// Restore the frosted-glass effect on the panels.
const glassComponents = [
    '.ucard', '.pivot-wrap', '.dash-tabs', '.ctrl-bar',
    '.pivot-banner-team', '.tk-dropdown', '.dash-sidebar',
];

for (const comp of glassComponents) {
    const escaped = comp.replace('.', '\\.');
    dash.subUnless(
        `${comp} backdrop blur`,
        // Skip if this component already has a blur, so re-running the script
        // doesn't stack a second backdrop-filter declaration on every pass.
        new RegExp(`${escaped}\\s*\\{[^}]*backdrop-filter`),
        new RegExp(`(${escaped}\\s*\\{[^}]*)background:\\s*var\\(--surf\\);`),
        '$1background: var(--surf);\n  backdrop-filter: blur(16px);\n  -webkit-backdrop-filter: blur(16px);'
    );
}

dash.sub(
    'dash header -> translucent',
    /\.dash-hdr\{[\s\S]*?\}/,
    `.dash-hdr{
  display:flex;justify-content:space-between;align-items:center;
  height:58px;padding:0 28px;
  background: rgba(10, 15, 30, 0.85);
  backdrop-filter: blur(16px);
  -webkit-backdrop-filter: blur(16px);
  border-bottom: 1px solid var(--bdr);
  position:sticky;top:0;z-index:200;
  box-shadow: 0 4px 20px rgba(0,0,0,0.3);
}`
);

dash.sub(
    'card headers -> blended',
    /\.ucard-hdr\{[\s\S]*?\}/,
    `.ucard-hdr{
  display:flex;justify-content:space-between;align-items:flex-start;
  padding:16px 16px 14px;
  background: rgba(255, 255, 255, 0.03);
  border-bottom: 1px solid var(--bdr);
}`
);

dash.sub(
    'pivot banner totals',
    /\.pivot-banner-total\s*\{[\s\S]*?\}/,
    `.pivot-banner-total {
  display:flex; flex-direction:column; align-items:center; justify-content:center;
  background: rgba(255, 255, 255, 0.05);
  border-radius:var(--r);
  padding:16px 24px; min-width:130px;
  border:1px solid var(--bdr);
  box-shadow:var(--sh);
}`
);

// ── report & save ─────────────────────────────────────────────────────────
const clean = [index.report(), dash.report()].every(Boolean);
index.save();
dash.save();

if (clean) {
    console.log('\nSalient Premium Navy theme applied.');
} else {
    console.error('\nSome rules did not match — the CSS has probably changed shape.');
    console.error('Check the "NO MATCH" lines above; .bak files hold the previous version.');
    process.exitCode = 1;
}
