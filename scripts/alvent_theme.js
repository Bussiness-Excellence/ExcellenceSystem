/**
 * alvent_theme.js — apply the Alvent violet dark theme.
 *
 * FIXED: every substitution now reports whether it matched, the files are
 * backed up before writing, and the script is idempotent (safe to re-run).
 * The old version printed "Alvent Violet theme applied!" even when every
 * single regex had failed to match.
 */

const { CssEditor } = require('./css_utils');

console.log('Applying Alvent Violet theme...');

// ── src/index.css ─────────────────────────────────────────────────────────
const index = new CssEditor('src/index.css');

index.sub(
    'body background -> violet radial gradient',
    /background:\s*radial-gradient\([^;]+;\s*background-attachment:\s*fixed;/,
    'background: #090514; ' +
    'background-image: radial-gradient(circle at 15% 50%, rgba(109, 40, 217, 0.15), transparent 40%), ' +
    'radial-gradient(circle at 85% 30%, rgba(139, 92, 246, 0.15), transparent 40%); ' +
    'background-attachment: fixed;'
);

// ── src/pages/Dashboard.css ───────────────────────────────────────────────
const dash = new CssEditor('src/pages/Dashboard.css');

dash.sub(
    'html.dark variables -> violet palette',
    /html\.dark\s*\{[\s\S]*?--sh3:[^\n]+\n\}/,
    `html.dark {
  --navy:   #f8fafc;
  --navy2:  #e2e8f0;
  --navy3:  #cbd5e1;
  --navy4:  #94a3b8;
  --gold:   #a78bfa; /* Violet accent */
  --gold2:  #c4b5fd;
  --gold-glow: rgba(139, 92, 246, 0.4);
  --bg:     transparent;
  --surf:   rgba(20, 10, 45, 0.6); /* Deep translucent violet */
  --bdr:    rgba(139, 92, 246, 0.2);
  --bdr2:   rgba(139, 92, 246, 0.35);
  --txt:    #ffffff;
  --mute:   #c4b5fd;
  --am:     #8b5cf6; /* Vibrant violet */
  --am-bg:  rgba(139, 92, 246, 0.2);
  --pm:     #6366f1; /* Indigo */
  --pm-bg:  rgba(99, 102, 241, 0.2);
  --sh: 0 4px 16px rgba(0, 0, 0, 0.4), inset 0 1px 1px rgba(139,92,246,0.15);
  --sh2: 0 8px 24px rgba(0, 0, 0, 0.5), inset 0 1px 2px rgba(139,92,246,0.25);
  --sh3: 0 12px 36px rgba(0, 0, 0, 0.6), inset 0 1px 3px rgba(139,92,246,0.3);
}`
);

dash.sub(
    'header background -> violet tint',
    /background:\s*rgba\(10, 15, 30, 0\.85\);/,
    'background: rgba(15, 8, 35, 0.85);'
);

dash.sub(
    'ucard background -> violet gradient',
    /\.ucard\s*\{[^}]*background:\s*var\(--surf\);[^}]*\}/g,
    match => match.replace(
        'background: var(--surf);',
        'background: linear-gradient(180deg, rgba(25, 12, 50, 0.7) 0%, rgba(45, 20, 85, 0.4) 100%);'
    )
);

dash.sub(
    'ucard header -> transparent with violet rule',
    /background:\s*rgba\(255, 255, 255, 0\.03\);/,
    'background: transparent; border-bottom: 1px solid rgba(139, 92, 246, 0.15);'
);

dash.subOnce(
    'ucard glowing top border',
    'border-top: 2px solid rgba(139, 92, 246, 0.6);',
    /\.ucard\s*\{/,
    '.ucard {\n  border-top: 2px solid rgba(139, 92, 246, 0.6);'
);

// ── report & save ─────────────────────────────────────────────────────────
const clean = [index.report(), dash.report()].every(Boolean);
index.save();
dash.save();

if (clean) {
    console.log('\nAlvent Violet theme applied.');
} else {
    console.error('\nSome rules did not match — the CSS has probably changed shape.');
    console.error('Check the "NO MATCH" lines above; .bak files hold the previous version.');
    process.exitCode = 1;
}
