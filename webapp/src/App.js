import React, { Suspense, lazy } from 'react';
import { HashRouter, Routes, Route, Navigate } from 'react-router-dom';
import { AuthProvider, useAuth } from './contexts/AuthContext';
import { ToastProvider } from './contexts/ToastContext';
import ProtectedRoute from './components/ProtectedRoute';
import Login from './components/Login';
import './App.css';

// Code-split the heavy routes. Previously every one of these was a static
// import, so a user landing on /login downloaded the 2,800-line Dashboard,
// the AdminPanel and the whole charting library before the login form could
// paint. Now each route's JS is fetched only when that route is opened.
//
// Login is deliberately NOT lazy: it is the first thing most visits render,
// and splitting it would add a network round trip to the critical path.
const Dashboard = lazy(() => import('./pages/Dashboard'));
const AdminPanel = lazy(() => import('./pages/AdminPanel'));
const ChangePassword = lazy(() => import('./components/ChangePassword'));

const Spinner = () => (
  <div className="app-loading"><div className="app-spinner" /></div>
);

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
  if (loading) return <Spinner />;
  if (session && profile) return <Navigate to="/dashboard" replace />;
  return <Login />;
}

function RootRoute() {
  const { session, profile, loading } = useAuth();
  if (loading) return <Spinner />;
  if (!session || !profile) return <Navigate to="/login" replace />;
  // All roles (including Admin) land on the dashboard by default
  return <Navigate to="/dashboard" replace />;
}

function AppContent() {
  const { session, profile, loading } = useAuth();

  if (loading) return <Spinner />;

  // If logged in but needs to change default password
  if (session && profile?.is_default_password) {
    return (
      <HashRouter>
        <Suspense fallback={<Spinner />}>
          <Routes>
            <Route path="*" element={<ChangePassword />} />
          </Routes>
        </Suspense>
      </HashRouter>
    );
  }

  return (
    <HashRouter>
      {/* One Suspense boundary around the whole route table: any lazy route
          shows the same spinner while its chunk downloads. */}
      <Suspense fallback={<Spinner />}>
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
      </Suspense>
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