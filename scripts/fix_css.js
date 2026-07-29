/**
 * fix_css.js — switch the default (light) theme to a pale surface with dark
 * text, keeping the dark palette behind html.dark.
 *
 * FIXED:
 *  - Every substitution now reports whether it matched. The old version wrote
 *    the file and printed "CSS updated successfully!" regardless, so a regex
 *    that no longer matched the CSS failed completely silently.
 *  - The file is backed up to Dashboard.css.bak before writing.
 *  - Three of the replacement patterns in the original list were exact
 *    duplicates of each other (the header, ctrl-bar and tabs rules shared an
 *    identical `linear-gradient(160deg, rgba(10,26,70,.80) ...)` pattern).
 *    Because .replace() with a /g regex consumes every occurrence on the
 *    first pass, the 2nd and 3rd entries could never match anything. They
 *    are now a single rule, so the "no match" warnings mean something.
 *  - The html.dark block is replaced with a non-greedy match anchored on
 *    --sidebar-bg, so it can't swallow following rules.
 */

const path = require('path');
const { CssEditor } = require('./css_utils');

const cssPath = path.join(__dirname, 'src/pages/Dashboard.css');
const editor = new CssEditor(cssPath);

// The goal is to make the default theme pale white and black text, and keep dark for html.dark.
// First, we replace root variables.
editor.sub('root variables -> pale light palette', /:root\s*\{[\s\S]*?\}/, `:root {
  --navy:   #1e293b;
  --navy2:  #334155;
  --navy3:  #475569;
  --navy4:  #64748b;
  --gold:   #2563eb;
  --gold2:  #3b82f6;
  --gold-glow: rgba(37, 99, 235, 0.15);
  
  --bg:    #f1f5f9;
  --surf:  #ffffff;
  --surf2: #f8fafc;
  --bdr:   #e2e8f0;
  --bdr2:  #cbd5e1;
  
  --txt:   #0f172a;
  --mute:  #64748b;
  
  --am:    #2563eb;
  --am-bg: rgba(37, 99, 235, 0.1);
  --pm:    #0891b2;
  --pm-bg: rgba(8, 145, 178, 0.1);
  
  --r:    14px;
  --r-sm: 10px;
  
  --sh:  0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06);
  --sh2: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06);
  --sh3: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -2px rgba(0,0,0,0.05);
  
  --grad-surf:    #ffffff;
  --grad-primary: linear-gradient(135deg, #1d4ed8 0%, #2563eb 50%, #3b82f6 100%);
  --grad-navy:    #f8fafc;
  --crystal-edge: transparent;
  
  --transition:   .2s cubic-bezier(0.4, 0, 0.2, 1);
  --glass-blur:   0px;
  
  --panel-bg: #ffffff;
  --panel-border: #e2e8f0;
  
  /* Sidebar specifically requested to be a bit darker than pale white */
  --sidebar-bg: #e2e8f0;
}`);

// Now replace hardcoded dark colors with var(--panel-bg), var(--panel-border), etc.
// Each entry is [label, pattern, replacement] so a failure can be reported.
const replacements = [
  // Header
  ['header background', /background:\s*linear-gradient\(160deg,\s*rgba\(10,26,70,\.90\)\s*0%,\s*rgba\(8,18,55,\.85\)\s*100%\);/g, 'background: var(--panel-bg);'],
  ['header bottom border', /border-bottom:\s*1px solid rgba\(77,159,255,\.22\);/g, 'border-bottom: 1px solid var(--panel-border);'],
  ['header shadow', /box-shadow:\s*0 1px 0 rgba\(77,159,255,\.12\),\s*0 4px 20px rgba\(0,5,30,\.40\);/g, 'box-shadow: var(--sh);'],

  // Ctrl bar + tabs share this gradient. The original list repeated this
  // pattern three times; with the /g flag the first pass already replaced
  // every occurrence, so entries 2 and 3 were dead code.
  ['ctrl-bar / tabs background', /background:\s*linear-gradient\(160deg,\s*rgba\(10,26,70,\.80\)\s*0%,\s*rgba\(8,18,55,\.75\)\s*100%\);/g, 'background: var(--panel-bg);'],
  ['ctrl-bar bottom border', /border-bottom:1px solid rgba\(77,159,255,\.16\);/g, 'border-bottom: 1px solid var(--panel-border);'],
  ['ctrl-bar shadow', /box-shadow:0 2px 12px rgba\(0,5,30,\.30\);/g, 'box-shadow: var(--sh);'],

  // Cards
  ['card background', /background:\s*linear-gradient\(160deg,\s*rgba\(14,34,80,\.72\)\s*0%,\s*rgba\(10,24,60,\.65\)\s*100%\);/g, 'background: var(--panel-bg);'],
  ['card border', /border:1px solid rgba\(77,159,255,\.18\);/g, 'border: 1px solid var(--panel-border);'],
  ['card shadow', /box-shadow:\s*0 2px 12px rgba\(0,8,40,\.40\),\s*0 1px 3px rgba\(0,8,40,\.30\);/g, 'box-shadow: var(--sh);'],

  // Pivot tables
  ['pivot background', /background:\s*linear-gradient\(160deg, #0a1c55 0%, #071232 100%\);/g, 'background: var(--panel-bg);'],
  ['pivot right border', /border-right: 1px solid rgba\(77,159,255,\.12\);/g, 'border-right: 1px solid var(--panel-border);'],

  // Sidebar
  ['sidebar background', /background:\s*linear-gradient\(160deg,\s*rgba\(10,26,70,\.88\)\s*0%,\s*rgba\(8,18,55,\.85\)\s*100%\);/g, 'background: var(--sidebar-bg);'],
  ['sidebar right border', /border-right:\s*1px solid rgba\(77,159,255,\.18\);/g, 'border-right: 1px solid var(--panel-border);'],

  // Text colours -> variables
  ['text #d4e8ff', /color:#d4e8ff/g, 'color:var(--txt)'],
  ['text #c8deff', /color:#c8deff/g, 'color:var(--txt)'],
  ['text #b3d4ff', /color:#b3d4ff/g, 'color:var(--txt)'],
  ['text #ffffff', /color:#ffffff/g, 'color:var(--txt)'],
  ['muted 177,210,255 .55', /color:rgba\(177,210,255,\.55\)/g, 'color:var(--mute)'],
  ['muted 150,190,255 .70', /color:rgba\(150,190,255,\.70\)/g, 'color:var(--mute)'],
  ['muted 150,190,255 .60', /color:rgba\(150,190,255,\.60\)/g, 'color:var(--mute)'],
  ['muted 177,210,255 .75', /color:rgba\(177,210,255,\.75\)/g, 'color:var(--mute)'],
  ['muted 177,210,255 .65', /color:rgba\(177,210,255,\.65\)/g, 'color:var(--mute)'],
  ['muted 150,190,255 .45', /color:rgba\(150,190,255,\.45\)/g, 'color:var(--mute)'],
  ['muted 150,190,255 .55', /color:rgba\(150,190,255,\.55\)/g, 'color:var(--mute)'],
  ['muted 150,190,255 .65', /color:rgba\(150,190,255,\.65\)/g, 'color:var(--mute)'],
  ['muted 120,165,220 .65', /color:rgba\(120,165,220,\.65\)/g, 'color:var(--mute)'],
  ['muted 126,200,255 .75', /color:rgba\(126,200,255,\.75\)/g, 'color:var(--mute)'],
  ['muted 160,195,240 .75', /color:rgba\(160,195,240,\.75\)/g, 'color:var(--mute)'],
  ['muted 191,207,255 .85', /color:rgba\(191,207,255,\.85\)/g, 'color:var(--mute)'],
];

replacements.forEach(([label, regex, replacement]) => {
  editor.sub(label, regex, replacement);
});

// For dark mode, we need to map --panel-bg, --panel-border, etc.
editor.sub('html.dark variables -> dark palette', /html\.dark\s*\{[\s\S]*?--sidebar-bg:[^\n]*\n\}/, `html.dark {
  --navy:   #f0f6ff;
  --navy2:  #d4e4ff;
  --navy3:  #a8c4f0;
  --navy4:  #7a9acc;
  --gold:   #60b8ff;
  --gold2:  #93d4ff;
  --gold-glow: rgba(96, 184, 255, 0.35);
  --surf:   rgba(8, 18, 48, 0.65);
  --surf2:  rgba(12, 25, 60, 0.75);
  --bdr:    rgba(60, 130, 246, 0.25);
  --bdr2:   rgba(96, 165, 250, 0.45);
  --txt:    #f4f8ff;
  --mute:   #7a9acc;
  --am:     #60b8ff;
  --am-bg:  rgba(96, 184, 255, 0.14);
  --pm:     #22d3ee;
  --pm-bg:  rgba(34, 211, 238, 0.14);
  --sh:  0 4px 20px rgba(0, 5, 20, 0.60);
  --sh2: 0 10px 32px rgba(0, 5, 20, 0.70);
  --sh3: 0 20px 56px rgba(0, 5, 20, 0.80);
  --grad-surf: linear-gradient(160deg, rgba(10,22,58,.80) 0%, rgba(6,14,38,.70) 100%);
  --crystal-edge: linear-gradient(120deg, rgba(96,184,255,.4) 0%, rgba(60,130,246,.06) 50%, rgba(96,184,255,.20) 100%);
  --panel-bg: rgba(14,34,80,.72);
  --panel-border: rgba(77,159,255,.18);
  --sidebar-bg: rgba(10,26,70,.88);
}`);

const clean = editor.report();
editor.save();

if (clean) {
  console.log('\nCSS updated successfully.');
} else {
  console.error('\nSome rules did not match — the CSS has probably changed shape.');
  console.error('Check the "NO MATCH" lines above; Dashboard.css.bak holds the previous version.');
  process.exitCode = 1;
}
