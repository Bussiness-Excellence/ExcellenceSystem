import { createClient } from '@supabase/supabase-js';

// These come from a .env file at the project root (never commit real
// values). Create react-app requires the REACT_APP_ prefix to expose
// variables to the browser bundle.
//
//   REACT_APP_SUPABASE_URL=https://xxbfwvlqixnmonxytdxq.supabase.co
//   REACT_APP_SUPABASE_ANON_KEY=your-anon-public-key
//
// IMPORTANT: this must be the anon/public key, never the service_role key.
// The service_role key bypasses Row Level Security entirely and must only
// ever be used in trusted server-side scripts (like the upload scripts),
// never shipped to the browser.
const supabaseUrl = process.env.REACT_APP_SUPABASE_URL;
const supabaseAnonKey = process.env.REACT_APP_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // Fail loudly in development rather than silently doing nothing.
  // eslint-disable-next-line no-console
  console.error(
    'Missing REACT_APP_SUPABASE_URL / REACT_APP_SUPABASE_ANON_KEY. ' +
    'Create a .env file at the project root — see supabaseClient.js for the format.'
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

// The internal fake-email domain used for username-style login.
// Must match EMAIL_DOMAIN in pulpoplus_provision_users.py.
export const EMAIL_DOMAIN = 'excellence-crm.internal';

export function employeeCodeToEmail(code) {
  return `${String(code).trim()}@${EMAIL_DOMAIN}`;
}
