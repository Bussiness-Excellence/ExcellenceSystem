const fs = require('fs');

const indexCssPath = 'src/index.css';
let indexCss = fs.readFileSync(indexCssPath, 'utf8');

// Update index.css background to the rich violet/purple gradient (Alvent theme)
indexCss = indexCss.replace(
  /background:\s*radial-gradient\([^;]+;\s*background-attachment:\s*fixed;/,
  'background: #090514; background-image: radial-gradient(circle at 15% 50%, rgba(109, 40, 217, 0.15), transparent 40%), radial-gradient(circle at 85% 30%, rgba(139, 92, 246, 0.15), transparent 40%); background-attachment: fixed;'
);

fs.writeFileSync(indexCssPath, indexCss);


const dashCssPath = 'src/pages/Dashboard.css';
let dashCss = fs.readFileSync(dashCssPath, 'utf8');

// Dashboard.css - Alvent Violet Theme
dashCss = dashCss.replace(
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

// Enhance headers with glowing violet borders to match the Alvent theme
dashCss = dashCss.replace(
  /background:\s*rgba\(10, 15, 30, 0\.85\);/,
  `background: rgba(15, 8, 35, 0.85);`
);

// ucard gradient background
dashCss = dashCss.replace(
  /\.ucard\s*\{[^}]*background:\s*var\(--surf\);[^}]*\}/g,
  match => match.replace('background: var(--surf);', 'background: linear-gradient(180deg, rgba(25, 12, 50, 0.7) 0%, rgba(45, 20, 85, 0.4) 100%);')
);

// ucard header transparent so gradient shows through
dashCss = dashCss.replace(
  /background:\s*rgba\(255, 255, 255, 0\.03\);/,
  `background: transparent; border-bottom: 1px solid rgba(139, 92, 246, 0.15);`
);

// Add a glowing top border to ucards to match the Alvent image style
if (!dashCss.includes('border-top: 2px solid rgba(139, 92, 246, 0.6);')) {
  dashCss = dashCss.replace(
    /\.ucard\s*\{/,
    `.ucard {\n  border-top: 2px solid rgba(139, 92, 246, 0.6);`
  );
}

fs.writeFileSync(dashCssPath, dashCss);

console.log("Alvent Violet theme applied!");
