import React from 'react';

/** Single shimmering skeleton block */
export function SkeletonBlock({ width = '100%', height = 16, radius = 8, style = {} }) {
  return (
    <div
      className="skeleton"
      style={{ width, height, borderRadius: radius, flexShrink: 0, ...style }}
    />
  );
}

/** A full skeleton KPI card that matches the real ucard layout */
export function SkeletonCard() {
  return (
    <div style={{
      background: 'rgba(30,41,59,0.4)',
      border: '1px solid rgba(255,255,255,0.06)',
      borderRadius: 16,
      overflow: 'hidden',
    }}>
      {/* Card header */}
      <div style={{
        padding: '20px',
        background: 'rgba(0,0,0,0.15)',
        borderBottom: '1px solid rgba(255,255,255,0.05)',
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'flex-start',
      }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8, flex: 1 }}>
          <SkeletonBlock width="60%" height={18} />
          <SkeletonBlock width="40%" height={13} />
          <SkeletonBlock width="70%" height={11} radius={6} style={{ marginTop: 4 }} />
        </div>
        <SkeletonBlock width={36} height={22} radius={6} />
      </div>

      {/* KPI rows */}
      {[1, 2, 3].map(s => (
        <div key={s} style={{ padding: '16px 20px', borderBottom: '1px solid rgba(255,255,255,0.04)' }}>
          <SkeletonBlock width="45%" height={10} style={{ marginBottom: 12 }} />
          {[1, 2, 3].map(r => (
            <div key={r} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '4px 0' }}>
              <SkeletonBlock width="38%" height={13} />
              <SkeletonBlock width={48} height={22} radius={6} />
            </div>
          ))}
        </div>
      ))}
    </div>
  );
}

/** Grid of skeleton cards */
export function SkeletonCardGrid({ count = 6 }) {
  return (
    <div style={{
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))',
      gap: 20,
    }}>
      {Array.from({ length: count }).map((_, i) => (
        <SkeletonCard key={i} />
      ))}
    </div>
  );
}

/** Skeleton for a pivot table */
export function SkeletonTable({ rows = 8, cols = 5 }) {
  return (
    <div style={{
      background: 'rgba(30,41,59,0.4)',
      border: '1px solid rgba(255,255,255,0.06)',
      borderRadius: 16,
      overflow: 'hidden',
    }}>
      {/* Header row */}
      <div style={{
        display: 'grid',
        gridTemplateColumns: `200px repeat(${cols - 1}, 1fr)`,
        gap: 1,
        background: 'rgba(15,23,42,0.6)',
        padding: '14px 16px',
        borderBottom: '1px solid rgba(255,255,255,0.06)',
      }}>
        {Array.from({ length: cols }).map((_, i) => (
          <SkeletonBlock key={i} width={i === 0 ? '70%' : '50%'} height={12} />
        ))}
      </div>
      {/* Data rows */}
      {Array.from({ length: rows }).map((_, ri) => (
        <div key={ri} style={{
          display: 'grid',
          gridTemplateColumns: `200px repeat(${cols - 1}, 1fr)`,
          gap: 1,
          padding: '12px 16px',
          borderBottom: '1px solid rgba(255,255,255,0.03)',
        }}>
          {Array.from({ length: cols }).map((_, ci) => (
            <SkeletonBlock key={ci} width={ci === 0 ? '80%' : `${Math.random() * 30 + 30}%`} height={13} />
          ))}
        </div>
      ))}
    </div>
  );
}
