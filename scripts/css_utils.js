/**
 * css_utils.js — shared helpers for the theme scripts.
 *
 * The original theme scripts ran a series of regex replacements and then
 * wrote the file and printed "applied successfully!" no matter what.
 * String.replace() returns the input unchanged when nothing matches, so a
 * regex that had drifted out of sync with the CSS failed completely
 * silently — the file was still rewritten, the success message still
 * printed, and the theme simply didn't change.
 *
 * These helpers make every substitution report whether it actually matched,
 * back up the file before writing, and skip the write entirely if nothing
 * changed (so the scripts are safe to re-run).
 */

const fs = require('fs');
const path = require('path');

class CssEditor {
    constructor(filePath) {
        this.path = filePath;
        this.ok = fs.existsSync(filePath);
        this.original = this.ok ? fs.readFileSync(filePath, 'utf8') : '';
        this.text = this.original;
        this.applied = [];
        this.missed = [];
    }

    /** Apply a replacement and record whether it matched. */
    sub(label, pattern, replacement) {
        if (!this.ok) return this;
        const before = this.text;
        this.text = this.text.replace(pattern, replacement);
        (this.text === before ? this.missed : this.applied).push(label);
        return this;
    }

    /** Apply only if `marker` is absent, so re-runs don't stack changes. */
    subOnce(label, marker, pattern, replacement) {
        if (!this.ok) return this;
        if (this.text.includes(marker)) {
            this.applied.push(`${label} (already present)`);
            return this;
        }
        return this.sub(label, pattern, replacement);
    }

    /**
     * Apply only if `guard` does NOT already match — for edits that would
     * otherwise stack a second copy of a declaration on every re-run.
     */
    subUnless(label, guard, pattern, replacement) {
        if (!this.ok) return this;
        if (guard.test(this.text)) {
            this.applied.push(`${label} (already present)`);
            return this;
        }
        return this.sub(label, pattern, replacement);
    }

    /** Write the file, backing up the previous contents first. */
    save() {
        if (!this.ok) {
            console.error(`  File not found: ${this.path}`);
            return false;
        }
        if (this.text === this.original) {
            console.log(`  ${path.basename(this.path)}: no changes needed`);
            return true;
        }
        const backup = `${this.path}.bak`;
        fs.writeFileSync(backup, this.original, 'utf8');
        fs.writeFileSync(this.path, this.text, 'utf8');
        console.log(`  ${path.basename(this.path)}: updated (backup at ${path.basename(backup)})`);
        return true;
    }

    report() {
        console.log(`\n${path.basename(this.path)}`);
        if (!this.ok) {
            console.error(`  MISSING — expected at ${this.path}`);
            console.error(`  Run this script from the webapp folder.`);
            return false;
        }
        this.applied.forEach(l => console.log(`  applied: ${l}`));
        this.missed.forEach(l => console.warn(`  NO MATCH: ${l}`));
        return this.missed.length === 0;
    }
}

module.exports = { CssEditor };
