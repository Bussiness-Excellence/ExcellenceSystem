/**
 * config.js — single source of truth for Supabase credentials (Node scripts).
 *
 * Loads from a .env file (see .env.example) and FAILS LOUDLY if anything is
 * missing, so a script can never silently run against the wrong project or
 * fall back to a key baked into the source.
 *
 * Usage:
 *     const { getServiceClient } = require('./config');
 *     const supabase = getServiceClient();
 */

const path = require('path');
const fs = require('fs');

// Add webapp/node_modules to module paths so dependencies are resolved
const webappNodeModules = path.join(__dirname, '..', 'webapp', 'node_modules');
if (fs.existsSync(webappNodeModules) && !module.paths.includes(webappNodeModules)) {
    module.paths.push(webappNodeModules);
}

// Zero-dependency .env loader (falls back gracefully if dotenv is missing)
const candidates = [
    path.join(__dirname, '.env'),
    path.join(__dirname, '..', '.env'),
];
for (const p of candidates) {
    if (fs.existsSync(p)) {
        try {
            require('dotenv').config({ path: p });
        } catch (e) {
            // Manual parse fallback if dotenv module is not installed
            const content = fs.readFileSync(p, 'utf8');
            content.split(/\r?\n/).forEach(line => {
                line = line.trim();
                if (!line || line.startsWith('#') || !line.includes('=')) return;
                const idx = line.indexOf('=');
                const k = line.slice(0, idx).trim();
                let v = line.slice(idx + 1).trim();
                if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) {
                    v = v.slice(1, -1);
                }
                if (!process.env[k]) {
                    process.env[k] = v;
                }
            });
        }
        break;
    }
}

function loadSupabaseModule() {
    try {
        return require('@supabase/supabase-js');
    } catch (e) {
        return require(path.join(__dirname, '..', 'webapp', 'node_modules', '@supabase', 'supabase-js'));
    }
}

function required(name, hint) {
    const v = process.env[name];
    if (!v || !v.trim()) {
        console.error(`\nMissing required environment variable: ${name}`);
        if (hint) console.error(`  ${hint}`);
        console.error(`  Copy .env.example to .env and fill it in.\n`);
        process.exit(1);
    }
    return v.trim();
}

/** Project URL, no trailing slash. */
function getUrl() {
    return required('SUPABASE_URL', 'Supabase dashboard -> Settings -> API -> Project URL')
        .replace(/\/+$/, '');
}

/**
 * Service-role key. Bypasses RLS — only for trusted local scripts,
 * never for anything that reaches a browser.
 */
function getServiceKey() {
    const key = required(
        'SUPABASE_SERVICE_ROLE_KEY',
        'Supabase dashboard -> Settings -> API -> service_role (secret)'
    );
    // Guard against someone pasting the public key into the secret slot:
    // writes would then fail confusingly at the RLS layer instead of here.
    if (key.startsWith('sb_publishable_') || /"role"\s*:\s*"anon"/.test(safeJwtPayload(key))) {
        console.error(
            '\nSUPABASE_SERVICE_ROLE_KEY looks like an anon/publishable key.' +
            '\nWrites and deletes will be blocked by Row Level Security.\n'
        );
        process.exit(1);
    }
    return key;
}

/** Anon key — safe for read-only inspection. */
function getAnonKey() {
    return required('SUPABASE_ANON_KEY', 'Supabase dashboard -> Settings -> API -> anon (public)');
}

function safeJwtPayload(token) {
    try {
        return Buffer.from(token.split('.')[1], 'base64').toString('utf8');
    } catch {
        return '';
    }
}

function getServiceClient() {
    const { createClient } = loadSupabaseModule();
    return createClient(getUrl(), getServiceKey(), {
        auth: { persistSession: false, autoRefreshToken: false },
    });
}

function getAnonClient() {
    const { createClient } = loadSupabaseModule();
    return createClient(getUrl(), getAnonKey(), {
        auth: { persistSession: false, autoRefreshToken: false },
    });
}

/** Resolve a path from env with a fallback, so drive layouts stay configurable. */
function envPath(name, fallback) {
    return (process.env[name] || fallback || '').trim();
}

module.exports = { getUrl, getServiceKey, getAnonKey, getServiceClient, getAnonClient, envPath };
