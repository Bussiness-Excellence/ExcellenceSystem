const { createClient } = require('@supabase/supabase-js');

module.exports = async function handler(req, res) {
  // Only allow POST requests
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  const { action, payload, token } = req.body;

  if (!token) {
    return res.status(401).json({ error: 'Authentication token required' });
  }

  try {
    const supabaseUrl = process.env.REACT_APP_SUPABASE_URL || 'https://xxbfwvlqixnmonxytdxq.supabase.co';
    const supabaseServiceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
    
    if (!supabaseServiceKey) {
      return res.status(500).json({ error: 'SUPABASE_SERVICE_ROLE_KEY is missing in Vercel Environment Variables. Please add it and redeploy.' });
    }

    const supabaseAdmin = createClient(supabaseUrl, supabaseServiceKey, {
      auth: { autoRefreshToken: false, persistSession: false }
    });
    // 1. Verify the caller's identity and role using their JWT token
    const { data: { user }, error: authError } = await supabaseAdmin.auth.getUser(token);
    if (authError || !user) {
      return res.status(401).json({ error: 'Invalid token' });
    }

    const { data: appUser } = await supabaseAdmin.from('app_users').select('role').eq('id', user.id).single();
    if (appUser?.role !== 'Admin') {
      return res.status(403).json({ error: 'Forbidden: Only Admins can perform this action' });
    }

    // 2. Perform the requested admin action
    if (action === 'reset_password') {
      const { target_user_id, new_password } = payload;
      const { error } = await supabaseAdmin.auth.admin.updateUserById(target_user_id, { password: new_password });
      if (error) throw error;
      return res.status(200).json({ success: true });
    }

    if (action === 'create_user') {
      const { email, password, code, name, role, teams } = payload;
      
      // Create Auth User
      const { data: authData, error: createAuthErr } = await supabaseAdmin.auth.admin.createUser({
        email,
        password,
        email_confirm: true
      });
      if (createAuthErr) throw createAuthErr;

      // Create app_user record
      const { error: dbErr } = await supabaseAdmin.from('app_users').insert({
        id: authData.user.id,
        employee_code: code,
        employee_name: name,
        role: role,
        visible_teams: role === 'Stakeholder' ? teams : null,
        is_active: true,
        is_default_password: true
      });

      if (dbErr) {
        await supabaseAdmin.auth.admin.deleteUser(authData.user.id);
        throw dbErr;
      }

      return res.status(200).json({ success: true });
    }

    return res.status(400).json({ error: 'Invalid action' });

  } catch (err) {
    console.error('Admin Action Error:', err);
    return res.status(500).json({ error: err.message || 'Internal Server Error' });
  }
}
