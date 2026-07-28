import React from 'react';
import { Navigate } from 'react-router-dom';
import { useAuth } from '../contexts/AuthContext';

export default function ProtectedRoute({ children, adminOnly = false }) {
  const { session, profile, loading } = useAuth();
  if (loading) return <div className="app-loading"><div className="app-spinner"/></div>;
  if (!session) return <Navigate to="/login" replace />;
  if (adminOnly && (profile?.role !== 'Admin' || ['4321', '5607'].includes(String(profile?.employee_code)))) return <Navigate to="/dashboard" replace />;
  return children;
}
