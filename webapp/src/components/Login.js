import React, { useState } from 'react';
import { useAuth } from '../contexts/AuthContext';
import './Login.css';

export default function Login() {
  const { signInWithEmployeeCode } = useAuth();
  const [employeeCode, setEmployeeCode] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    setError('');
    setSubmitting(true);
    const { error } = await signInWithEmployeeCode(employeeCode.trim(), password);
    setSubmitting(false);
    if (error) {
      setError('Employee code or password is incorrect.');
    }
  }

  return (
    <div className="login-screen">
      {/* Ambient mid orb */}
      <div className="login-orb-mid" />

      <div className="login-card">
        {/* Logo pill */}
        <div className="login-logos" style={{ justifyContent: 'center' }}>
          <div className="login-logos-pill" style={{ background: 'transparent', border: 'none', boxShadow: 'none' }}>
            <img src="/eipico_logo.png" alt="EIPICO Excellence" className="login-logo-company" style={{ height: '90px', width: 'auto', filter: 'drop-shadow(0 0 15px rgba(255,255,255,0.6))' }} />
          </div>
        </div>

        <h1>Excellence-CRM</h1>
        <p className="login-sub">Sign in with your employee credentials.</p>

        <form onSubmit={handleSubmit}>
          {/* Employee code */}
          <div className="login-field">
            <label htmlFor="employeeCode">Employee Code</label>
            <div className="login-input-wrap">
              <span className="login-input-icon">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                  <circle cx="12" cy="7" r="4"/>
                </svg>
              </span>
              <input
                id="employeeCode"
                type="text"
                inputMode="numeric"
                autoComplete="username"
                placeholder="e.g. 12345"
                value={employeeCode}
                onChange={(e) => setEmployeeCode(e.target.value)}
                required
                autoFocus
              />
            </div>
          </div>

          {/* Password */}
          <div className="login-field">
            <label htmlFor="password">Password</label>
            <div className="login-input-wrap">
              <span className="login-input-icon">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
                  <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
                </svg>
              </span>
              <input
                id="password"
                type="password"
                autoComplete="current-password"
                placeholder="••••••••"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                required
              />
            </div>
          </div>

          {error && (
            <div className="login-error" role="alert">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
              </svg>
              {error}
            </div>
          )}

          <button type="submit" className="login-submit" disabled={submitting}>
            <span className="login-btn-inner">
              {submitting && <span className="login-btn-spinner" />}
              {submitting ? 'Signing in…' : 'Sign In'}
            </span>
          </button>
        </form>

        <div className="login-divider">secured with supabase</div>

        <p className="login-hint">
          First time? Your password is your <strong>employee code</strong>.
          You'll be prompted to set a new one.
        </p>
      </div>
    </div>
  );
}