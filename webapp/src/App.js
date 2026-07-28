import React from 'react';
import { HashRouter, Routes, Route, Navigate } from 'react-router-dom';
import { AuthProvider, useAuth } from './contexts/AuthContext';
import { ToastProvider } from './contexts/ToastContext';
import ProtectedRoute from './components/ProtectedRoute';
import Login from './components/Login';
import Dashboard from './pages/Dashboard';
import AdminPanel from './pages/AdminPanel';
import ChangePassword from './components/ChangePassword';

import './App.css';

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null, errorInfo: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  componentDidCatch(error, errorInfo) {
    console.error("Uncaught runtime error:", error, errorInfo);
    this.setState({ errorInfo });
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: '40px', color: '#fff', backgroundColor: '#0f172a', minHeight: '100vh', fontFamily: 'sans-serif' }}>
          <h2 style={{ color: '#ef4444', marginBottom: '16px' }}>⚠️ Dashboard Diagnostics</h2>
          <div style={{ background: '#1e293b', padding: '20px', borderRadius: '8px', overflow: 'auto', border: '1px solid #334155' }}>
            <p style={{ fontWeight: 'bold', color: '#f87171', fontSize: '16px' }}>{this.state.error && this.state.error.toString()}</p>
            <pre style={{ color: '#94a3b8', fontSize: '12px', marginTop: '10px', whiteSpace: 'pre-wrap' }}>
              {this.state.errorInfo && this.state.errorInfo.componentStack}
            </pre>
          </div>
          <button 
            onClick={() => { sessionStorage.clear(); window.location.reload(); }} 
            style={{ marginTop: '20px', padding: '12px 24px', backgroundColor: '#3b82f6', color: '#fff', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
            🔄 Reset Cache & Reload App
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

function LoginRoute() {
  const { session, profile, loading } = useAuth();
  if (loading) return <div className="app-loading"><div className="app-spinner"/></div>;
  if (session && profile) return <Navigate to="/dashboard" replace />;
  return <Login />;
}

function RootRoute() {
  const { session, profile, loading } = useAuth();
  if (loading) return <div className="app-loading"><div className="app-spinner"/></div>;
  if (!session || !profile) return <Navigate to="/login" replace />;
  // All roles (including Admin) land on the dashboard by default
  return <Navigate to="/dashboard" replace />;
}

function AppContent() {
  const { session, profile, loading } = useAuth();
  
  if (loading) {
    return (
      <div className="app-loading" style={{display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100vh'}}>
        <div className="app-spinner" style={{
          width: '40px', height: '40px', borderRadius: '50%', border: '3px solid rgba(0,0,0,0.1)', borderTopColor: '#3b82f6', animation: 'spin 1s linear infinite'
        }}/>
      </div>
    );
  }

  // If logged in but needs to change default password
  if (session && profile?.is_default_password) {
    return (
      <HashRouter>
        <Routes>
          <Route path="*" element={<ChangePassword />} />
        </Routes>
      </HashRouter>
    );
  }

  return (
    <HashRouter>
      <Routes>
        <Route path="/login" element={<LoginRoute />} />
        <Route path="/dashboard" element={
          <ProtectedRoute>
            <Dashboard />
          </ProtectedRoute>
        } />
        <Route path="/admin" element={
          <ProtectedRoute adminOnly={true}>
            <AdminPanel />
          </ProtectedRoute>
        } />
        <Route path="/" element={<RootRoute />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </HashRouter>
  );
}

export default function App() {
  return (
    <ErrorBoundary>
      <ToastProvider>
        <AuthProvider>
          <AppContent />
        </AuthProvider>
      </ToastProvider>
    </ErrorBoundary>
  );
}
