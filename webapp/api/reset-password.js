const { createClient } = require('@supabase/supabase-js');

module.exports = async (req, res) => {
  // Allow CORS
  res.setHeader('Access-Control-Allow-Credentials', true);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version, Authorization'
  );

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { user_id, password } = req.body;
  if (!user_id || !password) {
    return res.status(400).json({ error: 'Missing user_id or password' });
  }

  const supabaseUrl = process.env.REACT_APP_SUPABASE_URL || 'https://xxbfwvlqixnmonxytdxq.supabase.co';
  const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh4YmZ3dmxxaXhubW9ueHl0ZHhxIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4Mjc1NjE2NSwiZXhwIjoyMDk4MzMyMTY1fQ.PSk6RyFmg_OFTcCtYO74AeJj6wT4FGZS2K2JT9GEJ_A)';

  const supabaseAdmin = createClient(supabaseUrl, serviceRoleKey, {
    auth: { autoRefreshToken: false, persistSession: false }
  });

  // Verify that the caller is an Admin
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'Unauthorized: Missing token' });
  }
  const token = authHeader.split(' ')[1];

  // Decode the user token to verify caller
  const { data: { user }, error: authErr } = await supabaseAdmin.auth.getUser(token);
  if (authErr || !user) {
    return res.status(401).json({ error: 'Unauthorized: Invalid token' });
  }

  // Check if this user is active and has the 'Admin' role in app_users
  const { data: callerProfile, error: profileErr } = await supabaseAdmin
    .from('app_users')
    .select('role, is_active')
    .eq('id', user.id)
    .single();

  if (profileErr || !callerProfile || !callerProfile.is_active || callerProfile.role !== 'Admin') {
    return res.status(403).json({ error: 'Forbidden: Admin access required' });
  }

  // Perform the password reset using Admin API
  const { error: resetErr } = await supabaseAdmin.auth.admin.updateUserById(user_id, {
    password: password
  });

  if (resetErr) {
    return res.status(500).json({ error: resetErr.message });
  }

  // Also reset the default password flag in app_users
  await supabaseAdmin
    .from('app_users')
    .update({ is_default_password: false })
    .eq('id', user_id);

  return res.status(200).json({ success: true, message: 'Password updated successfully' });
};
