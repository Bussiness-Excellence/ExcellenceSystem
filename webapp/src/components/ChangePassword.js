import React, { useState } from 'react';
import { supabase } from '../supabaseClient';
import { useAuth } from '../contexts/AuthContext';
import './Login.css'; // reuses the same visual language as the login screen

export default function ChangePassword() {
  const { profile, session } = useAuth();
  const [newPassword, setNewPassword] = useState('');
  const [confirm, setConfirm] = useState('');
  const [error, setError] = useState('');
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    setError('');

    if (newPassword.length < 8) {
      setError('Choose a password with at least 8 characters.');
      return;
    }
    if (newPassword !== confirm) {
      setError("Passwords don't match.");
      return;
    }

    setSubmitting(true);
    const { error: updateErr } = await supabase.auth.updateUser({ password: newPassword });
    if (updateErr) {
      setSubmitting(false);
      setError(updateErr.message);
      return;
    }

    const { error: flagErr } = await supabase
      .from('app_users')
      .update({ is_default_password: false })
      .eq('id', session.user.id);

    setSubmitting(false);
    if (flagErr) {
      // The password itself changed successfully; this is a secondary
      // bookkeeping update, so surface it but don't block the user.
      // eslint-disable-next-line no-console
      console.error('Password changed, but failed to clear the reset flag:', flagErr.message);
    }
    // AuthContext's onAuthStateChange listener doesn't refire on a password
    // update, so reload to pick up the cleared is_default_password flag.
    window.location.reload();
  }

  return (
    <div className="login-screen">
      <div className="login-card">
        <div className="login-logos">
          <div className="login-logos-pill">
            <img src="/eipico-logo.png" alt="EIPICO" className="login-logo-company" />
            <div className="login-logos-divider" />
            <img src="/dept-logo.png" alt="Excellence Department" className="login-logo-dept" />
          </div>
        </div>
        <h1>Set a new password</h1>
        <p className="login-sub">
          {profile ? `Welcome, ${profile.employee_name}. ` : ''}
          This is your first sign-in — choose a password only you know.
        </p>

        <form onSubmit={handleSubmit}>
          <label htmlFor="newPassword">New password</label>
          <input
            id="newPassword"
            type="password"
            value={newPassword}
            onChange={(e) => setNewPassword(e.target.value)}
            required
            autoFocus
          />

          <label htmlFor="confirm">Confirm password</label>
          <input
            id="confirm"
            type="password"
            value={confirm}
            onChange={(e) => setConfirm(e.target.value)}
            required
          />

          {error && <div className="login-error" role="alert">{error}</div>}

          <button type="submit" disabled={submitting}>
            {submitting ? 'Saving…' : 'Save and continue'}
          </button>
        </form>
      </div>
    </div>
  );
}
