require('dotenv').config();
const { createClient } = require('@supabase/supabase-js');
const s = createClient(process.env.REACT_APP_SUPABASE_URL, process.env.REACT_APP_SUPABASE_ANON_KEY);

async function check() {
  const { data, error } = await s.from('visits').select('*').eq('user', 'Maher Abd Elfattah').eq('visit_date', '2026-07-01');
  console.log("Maher visits on July 1st:", data?.length, data);
  
  const { data: d2 } = await s.from('visits').select('*').eq('user', 'Abdullah ahmed mohamed Abdelaal').eq('visit_date', '2026-07-01');
  console.log("Abdullah visits on July 1st:", d2?.length, d2);
}
check();
