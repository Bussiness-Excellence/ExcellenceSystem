import React, { useEffect, useMemo, useState, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../supabaseClient';
import { useAuth } from '../contexts/AuthContext';
import './Dashboard.css';

// ── i18n ──────────────────────────────────────────────────────────────────────
const T = {
  en: {
    brand: 'EXCELLENCE', signOut: 'Sign out', adminPanel: 'Admin Panel',
    lastMonth: 'Prev. Month Data', recent: 'Recent Month Data',
    allTeams: 'All teams', allUsers: 'All reps', search: 'Search name or territory…',
    export: 'Export', loading: 'Loading…', noData: 'No data for this period.',
    shiftAll: 'Both', shiftAM: 'AM', shiftPM: 'PM',
    people: n => `${n} rep${n !== 1 ? 's' : ''}`,
    tabs: { summary: 'Summary', specialty: 'Specialty', products: 'Products', coaching: 'Coaching/DV', timing: 'Last Visit Data' },
    roleView: { MR: 'My Results', Supervisor: 'My Team', 'Area Manager': 'My Area', BLM: 'Full Team', Admin: 'All Teams' },
    avg: 'Avg', sum: 'Sum', teamSummary: 'Team Summary',
    kpiGroups: [
      { label: 'Field Activity', keys: ['working_days', 'complete_field_days', 'am_shift_days', 'pm_shift_days', 'double_visit_days', 'office_work_days', 'no_activities', 'no_events'] },
      { label: 'Doctor Calls', keys: ['am_calls', 'am_call_rate', 'pm_calls', 'pm_call_rate'] },
      { label: 'Coverage', keys: ['total_am_covered', 'total_pm_covered', 'amcenter_covered', 'hospital_covered', 'am_accounts_unique', 'am_accounts_revisits', 'clinic_covered', 'polyclinic_covered'] },
      { label: 'Pharmacy', keys: ['pharmacies_visited', 'pharmacies_covered'] },
      { label: 'Products', keys: ['total_product_calls', 'distinct_products'] },
      { label: 'Coaching/DV', keys: ['coaching_days'] },
      { label: 'Last Visit Data', keys: ['avg_am_start_time', 'avg_am_shift_hm', 'avg_pm_shift_hm'] },
    ],
    kpi: {
      working_days: 'Working Days', complete_field_days: 'Field Days',
      am_shift_days: 'AM Days', pm_shift_days: 'PM Days',
      am_calls: 'AM Calls', pm_calls: 'PM Calls',
      am_call_rate: 'AM Call Rate', pm_call_rate: 'PM Call Rate',
      total_am_covered: 'AM Covered', total_pm_covered: 'PM Covered',
      amcenter_covered: 'AM Center', hospital_covered: 'Hospital',
      am_accounts_unique: 'AM Accounts', am_accounts_revisits: 'AM Revisits',
      clinic_covered: 'Clinic', polyclinic_covered: 'Poly Clinic',
      double_visit_days: 'Double Visits', coaching_days: 'Coaching/DV Days',
      office_work_days: 'Office Work', no_activities: 'Activities', no_events: 'Events',
      pharmacies_visited: 'Pharm. Visits', pharmacies_covered: 'Pharm. Covered',
      total_product_calls: 'Product Calls', distinct_products: 'Products',
      avg_am_start_time: 'AM Start Time', avg_am_shift_hm: 'AM Duration', avg_pm_shift_hm: 'PM Duration',
      timing_early: 'Before 3 PM', timing_normal: '3 PM – 6 PM', timing_late: 'After 6 PM',
    },
  },
  ar: {
    brand: 'إكسيلنس', signOut: 'خروج', adminPanel: 'لوحة الإدارة',
    lastMonth: 'الشهر الماضي', recent: 'الأحدث  1–15',
    allTeams: 'كل الفرق', allUsers: 'كل المندوبين', search: 'بحث باسم أو منطقة…',
    export: 'تصدير', loading: 'جارٍ التحميل…', noData: 'لا توجد بيانات.',
    shiftAll: 'الكل', shiftAM: 'AM', shiftPM: 'PM',
    people: n => `${n} مندوب`,
    tabs: { summary: 'الملخص', specialty: 'التخصص', products: 'المنتجات', coaching: 'التوجيه/مزدوجة', timing: 'بيانات الزيارة الأخيرة' },
    roleView: { MR: 'نتائجي', Supervisor: 'فريقي', 'Area Manager': 'منطقتي', BLM: 'الفريق', Admin: 'الكل' },
    avg: 'متوسط', sum: 'مجموع', teamSummary: 'ملخص الفريق',
    kpiGroups: [
      { label: 'النشاط الميداني', keys: ['working_days', 'complete_field_days', 'am_shift_days', 'pm_shift_days', 'double_visit_days', 'office_work_days', 'no_activities', 'no_events'] },
      { label: 'الزيارات', keys: ['am_calls', 'am_call_rate', 'pm_calls', 'pm_call_rate'] },
      { label: 'التغطية', keys: ['total_am_covered', 'total_pm_covered', 'amcenter_covered', 'hospital_covered', 'am_accounts_unique', 'am_accounts_revisits', 'clinic_covered', 'polyclinic_covered'] },
      { label: 'الصيدليات', keys: ['pharmacies_visited', 'pharmacies_covered'] },
      { label: 'المنتجات', keys: ['total_product_calls', 'distinct_products'] },
      { label: 'التوجيه/مزدوجة', keys: ['coaching_days'] },
      { label: 'بيانات الزيارة الأخيرة', keys: ['avg_am_start_time', 'avg_am_shift_hm', 'avg_pm_shift_hm'] },
    ],
    kpi: {
      working_days: 'أيام العمل', complete_field_days: 'أيام الميدان',
      am_shift_days: 'أيام AM', pm_shift_days: 'أيام PM',
      am_calls: 'زيارات AM', pm_calls: 'زيارات PM',
      am_call_rate: 'معدل AM', pm_call_rate: 'معدل PM',
      total_am_covered: 'تغطية AM', total_pm_covered: 'تغطية PM',
      amcenter_covered: 'مراكز AM', hospital_covered: 'مستشفيات',
      am_accounts_unique: 'حسابات AM', am_accounts_revisits: 'إعادة زيارات AM',
      clinic_covered: 'عيادات', polyclinic_covered: 'مراكز صحية',
      double_visit_days: 'زيارات مزدوجة', coaching_days: 'أيام التوجيه/مزدوجة',
      office_work_days: 'مكتب', no_activities: 'الأنشطة', no_events: 'الفعاليات',
      pharmacies_visited: 'زيارات صيدليات', pharmacies_covered: 'تغطية صيدليات',
      total_product_calls: 'مكالمات منتج', distinct_products: 'منتجات',
      avg_am_start_time: 'بدء AM', avg_am_shift_hm: 'مدة AM', avg_pm_shift_hm: 'مدة PM',
      timing_early: 'قبل 3 م', timing_normal: '3 م – 6 م', timing_late: 'بعد 6 م',
    },
  },
};

const NUMERIC_KPI_KEYS = [
  'working_days', 'complete_field_days', 'am_shift_days', 'pm_shift_days', 'double_visit_days', 'office_work_days',
  'no_activities', 'no_events',
  'am_calls', 'am_call_rate', 'pm_calls', 'pm_call_rate',
  'total_am_covered', 'total_pm_covered', 'amcenter_covered', 'hospital_covered', 'clinic_covered', 'polyclinic_covered',
  'pharmacies_visited', 'pharmacies_covered',
  'total_product_calls', 'distinct_products', 'coaching_days',
  'avg_am_shift_hm', 'avg_pm_shift_hm'
];

const PIE_COLORS = [
  '#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', '#ef4444',
  '#06b6d4', '#ec4899', '#84cc16', '#f97316', '#6366f1',
  '#14b8a6', '#e11d48', '#a855f7', '#0ea5e9', '#eab308',
];

// ── Helpers ──────────────────────────────────────────────────────────────────
function fmtDuration(decimalHours) {
  const h = Math.floor(decimalHours);
  const m = Math.round((decimalHours - h) * 60);
  if (h === 0 && m === 0) return '—';
  if (h === 0) return `${m}m`;
  if (m === 0) return `${h}h`;
  return `${h}h ${m}m`;
}

function fmtVal(v, key) {
  if (v === null || v === undefined || v === '') return '—';
  if (key === 'avg_am_shift_hm' || key === 'avg_pm_shift_hm') return fmtDuration(Number(v));
  if (key?.includes('rate')) return Number(v).toFixed(1);
  if (typeof v === 'number') return Number.isInteger(v) ? v : Number(v).toFixed(1);
  return v;
}

function sortSummary(rows) {
  return [...rows].sort((a, b) => {
    const tc = (a.team || '').localeCompare(b.team || '');
    if (tc) return tc;
    if (a.is_manager !== b.is_manager) return a.is_manager ? -1 : 1;
    return (a.user_name || '').localeCompare(b.user_name || '');
  });
}

function parseTimeToMinutes(timeStr) {
  if (!timeStr || typeof timeStr !== 'string') return null;
  const match = timeStr.match(/(\d{1,2}):(\d{2})(?:\s*(AM|PM))?/i);
  if (!match) return null;
  let hrs = parseInt(match[1], 10);
  const mins = parseInt(match[2], 10);
  const ampm = match[3] ? match[3].toUpperCase() : null;
  if (ampm === 'PM' && hrs < 12) hrs += 12;
  if (ampm === 'AM' && hrs === 12) hrs = 0;
  return hrs * 60 + mins;
}

function formatMinutesToTime(totalMins) {
  if (totalMins === null || isNaN(totalMins)) return '—';
  let hrs = Math.floor(totalMins / 60) % 24;
  const mins = Math.round(totalMins % 60);
  const ampm = hrs >= 12 ? 'PM' : 'AM';
  hrs = hrs % 12;
  if (hrs === 0) hrs = 12;
  const paddedMins = mins < 10 ? `0${mins}` : mins;
  return `${hrs}:${paddedMins} ${ampm}`;
}

function calcAvgStartTime(rows) {
  const minList = rows
    .map(r => parseTimeToMinutes(r.avg_am_start_time))
    .filter(m => m !== null);
  if (!minList.length) return '—';
  const avgMins = minList.reduce((a, b) => a + b, 0) / minList.length;
  return formatMinutesToTime(avgMins);
}

function computeAggregates(rows) {
  const reps = rows.filter(r => !r.is_manager);
  const targetRows = reps.length ? reps : rows;
  const agg = {};
  NUMERIC_KPI_KEYS.forEach(key => {
    const vals = targetRows.map(r => Number(r[key]) || 0).filter(v => v > 0);
    agg[key] = {
      sum: targetRows.map(r => Number(r[key]) || 0).reduce((s, v) => s + v, 0),
      avg: vals.length ? (vals.reduce((s, v) => s + v, 0) / vals.length) : 0,
    };
  });
  return { agg, repCount: reps.length || rows.length };
}

// ── PieChart (SVG donut) ─────────────────────────────────────────────────────
// ── KPI targets for progress indicators ─────────────────────────────────────
const KPI_TARGETS = {
  working_days: 22,
  complete_field_days: 20,
  am_calls: 120,
  pm_calls: 120,
  total_am_covered: 80,
  total_pm_covered: 80,
  pharmacies_visited: 40,
  coaching_days: 4,
};

// ── PieChart (SVG donut) ─────────────────────────────────────────────────────
function PieChart({ data, title, size = 140, thickness = 22, onSelect, activeFilters = new Set() }) {
  const [hoveredLabel, setHoveredLabel] = useState(null);
  const total = data.reduce((s, d) => s + d.value, 0);
  if (!total) return <div className="pie-empty">No data</div>;
  const center = size / 2;
  const radius = (size - thickness) / 2;
  const circumference = 2 * Math.PI * radius;
  let accumulated = 0;
  const segments = data.map((d) => {
    const pct = d.value / total;
    const arc = pct * circumference;
    const offset = accumulated;
    accumulated += arc;
    return { ...d, pct, arc, offset };
  });

  const hasSelections = activeFilters && activeFilters.size > 0;

  return (
    <div className="pie-chart">
      {title && <div className="pie-title">{title}</div>}
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} className="pie-svg animated-pie">
        <circle cx={center} cy={center} r={radius} fill="none" stroke="rgba(255,255,255,.06)" strokeWidth={thickness} />
        {segments.map((seg, i) => {
          const isSelected = activeFilters?.has(seg.label);
          const isHovered = hoveredLabel === seg.label;
          const fade = (hasSelections && !isSelected) || (hoveredLabel && !isHovered);
          const currentThickness = isHovered ? thickness + 6 : (isSelected ? thickness + 4 : thickness);
          return (
            <circle key={`${seg.label}-${i}`} cx={center} cy={center} r={radius} fill="none"
              stroke={seg.color} strokeWidth={currentThickness}
              strokeDasharray={`${seg.arc} ${circumference - seg.arc}`}
              strokeDashoffset={-seg.offset}
              transform={`rotate(-90 ${center} ${center})`}
              className={`pie-segment ${isHovered ? 'hovered' : ''}`} strokeLinecap="butt"
              style={{
                cursor: 'pointer',
                opacity: fade ? 0.25 : 1,
                transition: 'stroke-width 0.3s ease, filter 0.3s ease, opacity 0.3s ease, stroke-dasharray 0.5s ease-out, stroke-dashoffset 0.5s ease-out',
                filter: isHovered ? `drop-shadow(0 0 8px ${seg.color})` : 'none',
              }}
              onMouseEnter={() => setHoveredLabel(seg.label)}
              onMouseLeave={() => setHoveredLabel(null)}
              onClick={() => onSelect && onSelect(seg.label)}
            />
          );
        })}
        <text x={center} y={center - 6} textAnchor="middle" dominantBaseline="central" className="pie-center-val">{total}</text>
        <text x={center} y={center + 10} textAnchor="middle" dominantBaseline="central" className="pie-center-lbl">calls</text>
      </svg>
      <div className="pie-legend">
        {segments.slice(0, 6).map((seg, i) => {
          const isSelected = activeFilters?.has(seg.label);
          const isHovered = hoveredLabel === seg.label;
          const fade = (hasSelections && !isSelected) || (hoveredLabel && !isHovered);
          return (
            <div key={i}
              className={`pie-leg-item ${isSelected ? 'selected' : ''} ${isHovered ? 'hovered' : ''}`}
              style={{
                cursor: 'pointer',
                opacity: fade ? 0.4 : 1,
                background: isHovered ? 'rgba(200, 168, 75, 0.2)' : (isSelected ? 'rgba(200, 168, 75, 0.12)' : 'none'),
                border: isHovered ? '1px solid var(--gold)' : (isSelected ? '1px solid rgba(200, 168, 75, 0.3)' : '1px solid transparent'),
                padding: '4px 6px',
                borderRadius: '4px',
                transform: isHovered ? 'translateX(3px)' : 'none',
                transition: 'all 0.25s ease',
              }}
              onMouseEnter={() => setHoveredLabel(seg.label)}
              onMouseLeave={() => setHoveredLabel(null)}
              onClick={() => onSelect && onSelect(seg.label)}
            >
              <span className="pie-dot" style={{ background: seg.color, transform: isHovered ? 'scale(1.4)' : 'scale(1)', transition: 'transform 0.2s ease' }} />
              <span className="pie-leg-label">{seg.label}</span>
              <span className="pie-leg-val">{Math.round(seg.pct * 100)}%</span>
            </div>
          );
        })}
        {segments.length > 6 && <div className="pie-leg-more">+{segments.length - 6} more</div>}
      </div>
    </div>
  );
}

// ── TeamBriefCard ─────────────────────────────────────────────────────────────
function TeamBriefCard({ rows, teamLabel, rtl, t, shift, isMgr, onSelectTeam }) {
  const reps = rows.filter(r => !r.is_manager);
  const repCount = reps.length || rows.length;

  const { agg, avgStartTime, totalActivities, avgActivities, totalEvents, avgEvents } = useMemo(() => {
    const { agg } = computeAggregates(rows);
    const avgStartTime = calcAvgStartTime(rows);
    const totalActivities = rows.reduce((s, r) => s + (Number(r.no_activities) || 0), 0);
    const avgActivities = repCount ? (totalActivities / repCount).toFixed(1) : '0';
    const totalEvents = rows.reduce((s, r) => s + (Number(r.no_events) || 0), 0);
    const avgEvents = repCount ? (totalEvents / repCount).toFixed(1) : '0';
    return { agg, avgStartTime, totalActivities, avgActivities, totalEvents, avgEvents };
  }, [rows, repCount]);

  const amShiftDur = agg['avg_am_shift_hm']?.avg || 0;
  const pmShiftDur = agg['avg_pm_shift_hm']?.avg || 0;

  return (
    <div className="ucard team-brief-card" onClick={() => onSelectTeam && onSelectTeam(teamLabel)} style={{ cursor: 'pointer' }}>
      <div className="ucard-hdr">
        <div className="ucard-info">
          <div className="ucard-name">{teamLabel}</div>
          <div className="ucard-meta">{repCount} {rtl ? 'مندوب' : 'reps'}</div>
          {(amShiftDur > 0 || pmShiftDur > 0) && (
            <div className="ucard-dur">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><circle cx="12" cy="12" r="10" /><path d="M12 6v6l4 2" /></svg>
              {amShiftDur > 0 ? <span className="dur-am">AM {fmtDuration(amShiftDur)}</span> : null}
              {pmShiftDur > 0 ? <span className="dur-pm">PM {fmtDuration(pmShiftDur)}</span> : null}
            </div>
          )}
        </div>
        <span className="mgr-pip">{rtl ? 'فريق' : 'TEAM'}</span>
      </div>

      {t.kpiGroups.map(g => {
        const keys = g.keys.filter(k => {
          if (shift === 'AM') return !['pm_calls', 'pm_call_rate', 'pm_shift_days', 'total_pm_covered', 'clinic_covered', 'polyclinic_covered', 'avg_pm_shift_hm'].includes(k);
          if (shift === 'PM') return !['am_calls', 'am_call_rate', 'am_shift_days', 'total_am_covered', 'amcenter_covered', 'hospital_covered', 'avg_am_shift_hm', 'avg_am_start_time'].includes(k);
          return true;
        });
        if (g.keys.includes('coaching_days') && !isMgr) return null;

        return (
          <div key={g.label} className={`kpi-sec${g.keys.includes('avg_am_start_time') ? ' kpi-timing' : ''}`}>
            <div className="kpi-sec-hd">{g.label}</div>
            {keys.map(k => {
              let displayVal = '—';
              if (k === 'avg_am_start_time') {
                displayVal = avgStartTime;
              } else if (k === 'no_activities') {
                displayVal = `${totalActivities} (${avgActivities}/rep)`;
              } else if (k === 'no_events') {
                displayVal = `${totalEvents} (${avgEvents}/rep)`;
              } else {
                const item = agg[k];
                if (item) {
                  displayVal = fmtVal(item.avg, k);
                }
              }
              return (
                <div key={k} className="kpi-row-wrapper">
                  <div className="kpi-row">
                    <span className="kpi-lbl">{t.kpi[k] || k}</span>
                    <span className={`kpi-v ${k.includes('rate') ? 'rate' : ''}`}>{displayVal}</span>
                  </div>
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}

// ── PivotSummaryBanner ───────────────────────────────────────────────────────
function PivotSummaryBanner({ rows, valueKey, rowKey, shift, t, selectedTeam, onSelectTeam, userTeamMap }) {
  const filtered = useMemo(() => shift === 'all' ? rows : rows.filter(r => r.shift === shift), [rows, shift]);
  const byTeam = useMemo(() => {
    const m = {};
    if (userTeamMap) {
      Object.entries(userTeamMap).forEach(([userName, teamStr]) => {
        if (!teamStr || teamStr === 'Unknown') return;
        const tms = (typeof teamStr === 'string') ? teamStr.split('; ') : [teamStr];
        tms.forEach(team => {
          if (!m[team]) m[team] = { total: 0, users: new Set() };
          m[team].users.add(userName);
        });
      });
    }
    filtered.forEach(r => {
      const teamStr = (r.team && r.team !== 'Unknown') ? r.team : (userTeamMap && userTeamMap[r.user_name]) || 'Other';
      const tms = (typeof teamStr === 'string') ? teamStr.split('; ') : [teamStr];
      tms.forEach(team => {
        if (!m[team]) m[team] = { total: 0, users: new Set() };
        m[team].total += (r[valueKey] || 0);
        m[team].users.add(r.user_name);
      });
    });
    return m;
  }, [filtered, valueKey, userTeamMap]);
  const grandTotal = useMemo(() => filtered.reduce((s, r) => s + (r[valueKey] || 0), 0), [filtered, valueKey]);
  const allUsers = useMemo(() => {
    if (userTeamMap) return Object.keys(userTeamMap).length;
    return new Set(filtered.map(r => r.user_name)).size;
  }, [filtered, userTeamMap]);
  const teamList = Object.entries(byTeam).sort((a, b) => a[0].localeCompare(b[0]));
  if (!teamList.length) return null;
  return (
    <div className="pivot-banner">
      <div
        className={`pivot-banner-total ${selectedTeam === 'all' ? 'active' : ''}`}
        onClick={() => onSelectTeam && onSelectTeam('all')}
        style={{ cursor: onSelectTeam ? 'pointer' : 'default' }}
      >
        <span className="pb-label">Grand Total</span>
        <span className="pb-val">{grandTotal.toLocaleString()}</span>
        <span className="pb-sub">{allUsers} reps</span>
      </div>
      {teamList.map(([team, d]) => (
        <div key={team}
          className={`pivot-banner-team ${selectedTeam === team ? 'active' : ''}`}
          onClick={() => onSelectTeam && onSelectTeam(selectedTeam === team ? 'all' : team)}
          style={{ cursor: onSelectTeam ? 'pointer' : 'default' }}
        >
          <span className="pb-team">{team}</span>
          <span className="pb-val">{d.total.toLocaleString()}</span>
          <span className="pb-sub">{d.users.size} reps · avg {d.users.size ? Math.round(d.total / d.users.size) : 0}</span>
        </div>
      ))}
    </div>
  );
}

// ── ShiftToggle ──────────────────────────────────────────────────────────────
function ShiftToggle({ value, onChange, t }) {
  return (
    <div className="shift-toggle">
      {['all', 'AM', 'PM'].map(s => (
        <button key={s}
          className={`stoggle${value === s ? ' on' : ''} ${s === 'AM' ? 'am' : s === 'PM' ? 'pm' : ''}`}
          onClick={() => onChange(s)}>
          {s === 'all' ? t.shiftAll : s}
        </button>
      ))}
    </div>
  );
}

// ── PivotTable ───────────────────────────────────────────────────────────────
function PivotTable({ rows, rowKey, valueKey, shiftFilter, userFilter, searchFilter, lang, hideAvg }) {
  const filtered = useMemo(() => rows.filter(r => {
    if (shiftFilter !== 'all' && r.shift !== shiftFilter) return false;
    if (userFilter && userFilter !== 'all' && r.user_name !== userFilter) return false;
    if (searchFilter && !r[rowKey]?.toLowerCase().includes(searchFilter.toLowerCase())
      && !r.user_name?.toLowerCase().includes(searchFilter.toLowerCase())) return false;
    return true;
  }), [rows, rowKey, shiftFilter, userFilter, searchFilter]);
  const users = useMemo(() => [...new Set(filtered.map(r => r.user_name))].sort(), [filtered]);
  const rowKeys = useMemo(() => [...new Set(filtered.map(r => r[rowKey]))].sort(), [filtered, rowKey]);
  const cells = useMemo(() => {
    const c = {};
    filtered.forEach(r => {
      const k = r[rowKey]; if (!c[k]) c[k] = {};
      c[k][r.user_name] = (c[k][r.user_name] || 0) + (r[valueKey] || 0);
    });
    return c;
  }, [filtered, rowKey, valueKey]);
  const colTotals = useMemo(() => {
    const ct = {};
    users.forEach(u => ct[u] = filtered.filter(r => r.user_name === u).reduce((s, r) => s + (r[valueKey] || 0), 0));
    return ct;
  }, [users, filtered, valueKey]);

  // Synced top scrollbar — lets people scroll horizontally without hunting
  // for the scrollbar at the bottom of a long table.
  const topScrollRef = useRef(null);
  const wrapRef = useRef(null);
  const [tableWidth, setTableWidth] = useState(0);
  const syncingRef = useRef(false);

  useEffect(() => {
    const wrapEl = wrapRef.current;
    if (!wrapEl) return;
    const table = wrapEl.querySelector('table');
    if (!table) return;
    const update = () => setTableWidth(table.scrollWidth);
    update();
    const ro = new ResizeObserver(update);
    ro.observe(table);
    return () => ro.disconnect();
  }, [filtered, users, rowKeys]);

  const handleTopScroll = () => {
    if (syncingRef.current) { syncingRef.current = false; return; }
    if (!wrapRef.current || !topScrollRef.current) return;
    syncingRef.current = true;
    wrapRef.current.scrollLeft = topScrollRef.current.scrollLeft;
  };
  const handleWrapScroll = () => {
    if (syncingRef.current) { syncingRef.current = false; return; }
    if (!wrapRef.current || !topScrollRef.current) return;
    syncingRef.current = true;
    topScrollRef.current.scrollLeft = wrapRef.current.scrollLeft;
  };

  if (!filtered.length) return <div className="dash-empty">{lang === 'ar' ? 'لا توجد بيانات' : 'No data'}</div>;
  return (
    <>
      <div className="pivot-top-scroll" ref={topScrollRef} onScroll={handleTopScroll}>
        <div style={{ width: tableWidth, height: 1 }} />
      </div>
      <div className="pivot-wrap" ref={wrapRef} onScroll={handleWrapScroll}>
        <table className="pivot-tbl">
          <thead>
            <tr>
              <th className="s-col">{rowKey === 'specialty' ? (lang === 'ar' ? 'التخصص' : 'Specialty') : (lang === 'ar' ? 'المنتج' : 'Product')}</th>
              {users.map(u => <th key={u} title={u}>{u.split(' ').slice(0, 2).join(' ')}</th>)}
              <th className="t-col">Σ Total</th>
            </tr>
            {!hideAvg && (
              <tr className="avg-row">
                <th className="s-col avg-lbl">⌀ Avg / rep</th>
                {users.map(u => {
                  const uTotal = colTotals[u] || 0;
                  const uRows = rowKeys.filter(k => cells[k]?.[u]).length;
                  return <th key={u} className="avg-cell">{uRows > 0 ? Math.round(uTotal / uRows) : 0}</th>;
                })}
                <th className="t-col avg-cell">
                  {(() => {
                    const gt = filtered.reduce((s, r) => s + (r[valueKey] || 0), 0);
                    return users.length > 0 ? Math.round(gt / users.length) : 0;
                  })()}
                </th>
              </tr>
            )}
          </thead>
          <tbody>
            {rowKeys.map(k => {
              const rowTotal = users.reduce((s, u) => s + (cells[k]?.[u] || 0), 0);
              return (
                <tr key={k}>
                  <td className="s-col">{k}</td>
                  {users.map(u => {
                    const v = cells[k]?.[u];
                    return <td key={u} className={v ? 'has-v' : 'nil'}>{v || ''}</td>;
                  })}
                  <td className="t-col">{rowTotal}</td>
                </tr>
              );
            })}
            <tr className="tot-row">
              <td className="s-col">Σ Total</td>
              {users.map(u => <td key={u}>{colTotals[u] || 0}</td>)}
              <td className="t-col">{filtered.reduce((s, r) => s + (r[valueKey] || 0), 0)}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </>
  );
}

// ── Main Dashboard ────────────────────────────────────────────────────────────
export default function Dashboard() {
  const { profile, hierarchy, visibleCodes, signOut } = useAuth();
  const [lang, setLang] = useState(profile?.preferred_lang || 'en');
  const [period, setPeriod] = useState('');
  const [availablePeriods, setAvailablePeriods] = useState([]);
  const [team, setTeam] = useState('all');
  const [shift, setShift] = useState('all'); // Default is Both
  const [timeGrain, setTimeGrain] = useState('all'); // 'all' | 'biweekly1' | 'biweekly2' | 'week1' | 'week2' | 'week3' | 'week4'
  const [search, setSearch] = useState('');

  useEffect(() => {
    async function loadPeriods() {
      try {
        const { data } = await supabase.from('summaries')
          .select('period')
          .neq('period', 'Recent')
          .neq('period', 'Last Month')
          .neq('period', 'Recent Month Data')
          .neq('period', 'Prev. Month Data')
          .neq('period', 'الأحدث  1–15')
          .neq('period', 'الشهر الماضي');
        if (data && data.length > 0) {
          const validPeriods = data.map(r => r.period).filter(p => p && !p.toLowerCase().includes('recent') && !p.toLowerCase().includes('last month'));
          const uniq = [...new Set(validPeriods)].sort((a, b) => {
            const dateA = Date.parse(a);
            const dateB = Date.parse(b);
            if (!isNaN(dateA) && !isNaN(dateB)) return dateB - dateA;
            return b.localeCompare(a);
          });
          if (uniq.length > 0) {
            setAvailablePeriods(uniq);
            setPeriod(prev => (uniq.includes(prev) ? prev : uniq[0]));
          }
        }
      } catch (e) {
        console.error("Error loading periods:", e);
      }
    }
    loadPeriods();
  }, []);

  const getTimeGrainRatio = useCallback(() => {
    if (timeGrain === 'all') return 1.0;
    if (timeGrain === 'biweekly1' || timeGrain === 'biweekly2') return 0.50;
    if (timeGrain === 'week1' || timeGrain === 'week2' || timeGrain === 'week3') return 7 / 30;
    if (timeGrain === 'week4') return 9 / 30;
    return 1.0;
  }, [timeGrain]);

  const normalizeDateStr = useCallback((dStr) => {
    if (!dStr) return '';
    const s = String(dStr).trim();

    // ISO format: YYYY-MM-DD or YYYY/MM/DD
    let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) {
      const yr = m[1];
      const mo = String(m[2]).padStart(2, '0');
      const da = String(m[3]).padStart(2, '0');
      return `${yr}-${mo}-${da}`;
    }

    // Standard CRM format: DD/MM/YYYY or MM/DD/YYYY
    m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (m) {
      const p1 = parseInt(m[1], 10);
      const p2 = parseInt(m[2], 10);
      const yr = m[3];

      let day, month;
      if (p2 > 12) {
        // MM/DD/YYYY format: p1 is month, p2 is day (e.g. 07/18/2026)
        month = p1;
        day = p2;
      } else {
        // Standard DD/MM/YYYY format: p1 is day, p2 is month (e.g. 08/07/2026)
        day = p1;
        month = p2;
      }
      return `${yr}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }

    return s;
  }, []);

  const filterByTimeGrain = useCallback((rows) => {
    if (!rows || !rows.length || timeGrain === 'all') return rows;
    return rows.filter(r => {
      const rawDate = r.visit_date || r.coaching_date || r.date;
      if (!rawDate) return false;
      const isoDate = normalizeDateStr(rawDate);
      if (!isoDate) return false;

      const parts = isoDate.split('-');
      if (parts.length < 3) return false;
      const day = parseInt(parts[2], 10);
      if (isNaN(day)) return false;

      if (timeGrain === 'biweekly1') return day >= 1 && day <= 15;
      if (timeGrain === 'biweekly2') return day >= 16 && day <= 31;
      if (timeGrain === 'week1') return day >= 1 && day <= 7;
      if (timeGrain === 'week2') return day >= 8 && day <= 14;
      if (timeGrain === 'week3') return day >= 15 && day <= 21;
      if (timeGrain === 'week4') return day >= 22 && day <= 31;
      return true;
    });
  }, [timeGrain, normalizeDateStr]);
  const [userFilter, setUser] = useState('all');
  const [tab, setTab] = useState('summary');
  const [rawSummary, setSummary] = useState([]);
  const [rawSpecialty, setSpecialty] = useState([]);
  const [rawProducts, setProducts] = useState([]);
  const [rawCoaching, setCoaching] = useState([]);
  const [rawVisits, setVisits] = useState([]);
  const [teamsMap, setTeamsMap] = useState({});
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');

  // Sidebar states
  const [selectedRep, setSelectedRep] = useState(null);
  const [specialtyFilter, setSpecialtyFilter] = useState(new Set());
  const [productFilter, setProductFilter] = useState(new Set());
  const [classificationFilter, setClassificationFilter] = useState(new Set());
  const [selectedManager, setSelectedManager] = useState(null);
  const [lineManagerFilter, setLineManagerFilter] = useState('all');
  const [managerTerritoryFilter, setManagerTerritoryFilter] = useState('all');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [theme, setTheme] = useState(() => localStorage.getItem('theme') || 'light');

  const userHierarchyMap = useMemo(() => {
    const map = {};
    (hierarchy || []).forEach(h => {
      if (h.employee_name) {
        const k = h.employee_name.toLowerCase().trim();
        if (!map[k]) {
          map[k] = {
            role: h.role,
            blm_name: h.blm_name,
            territory: h.division_name,
            area_manager: h.area_manager_name,
            supervisor: h.supervisor_name,
          };
        }
      }
    });
    return map;
  }, [hierarchy]);

  const userTeamMap = useMemo(() => {
    const map = {};

    const addTeam = (mgrName, teamStr) => {
      if (!mgrName || !teamStr || teamStr === 'Unknown') return;
      const norm = mgrName.toLowerCase().trim();
      if (!map[norm]) map[norm] = new Set();
      teamStr.split(/;\s*/).filter(Boolean).forEach(t => map[norm].add(t));
    };

    // 1. Gather from rawSummary
    rawSummary.forEach(r => {
      if (r.user_name && r.team) addTeam(r.user_name, r.team);
    });

    // 2. Gather from complete hierarchy using teamsMap
    (hierarchy || []).forEach(h => {
      const teamName = teamsMap[h.team_id];
      if (!teamName) return;

      if (h.employee_name) addTeam(h.employee_name, teamName);
      if (h.supervisor_name) addTeam(h.supervisor_name, teamName);
      if (h.area_manager_name) addTeam(h.area_manager_name, teamName);
      if (h.blm_name && !h.blm_name.toLowerCase().includes('directory') && !h.blm_name.toLowerCase().includes('team')) {
        addTeam(h.blm_name, teamName);
      }
    });

    const finalMap = {};
    Object.keys(map).forEach(k => {
      finalMap[k] = Array.from(map[k]).sort().join('; ');
    });
    return finalMap;
  }, [rawSummary, hierarchy, teamsMap]);

  const summary = useMemo(() => rawSummary.map(r => ({ ...r, team: userTeamMap[r.user_name?.toLowerCase().trim()] || r.team })), [rawSummary, userTeamMap]);
  const specialty = useMemo(() => rawSpecialty.map(r => ({ ...r, team: userTeamMap[r.user_name?.toLowerCase().trim()] || r.team })), [rawSpecialty, userTeamMap]);
  const products = useMemo(() => rawProducts.map(r => ({ ...r, team: userTeamMap[r.user_name?.toLowerCase().trim()] || r.team })), [rawProducts, userTeamMap]);
  const coaching = useMemo(() => rawCoaching.map(r => ({ ...r, team: userTeamMap[r.manager_name?.toLowerCase().trim()] || r.team })), [rawCoaching, userTeamMap]);



  useEffect(() => {
    if (theme === 'dark') {
      document.documentElement.classList.add('dark');
      document.documentElement.classList.remove('light');
    } else {
      document.documentElement.classList.add('light');
      document.documentElement.classList.remove('dark');
    }
  }, [theme]);

  // AI Chat States
  const [isAiOpen, setIsAiOpen] = useState(false);
  const [aiInput, setAiInput] = useState('');
  const [aiHistory, setAiHistory] = useState([]);
  const [isAiLoading, setIsAiLoading] = useState(false);

  const toggleTheme = () => {
    const nextTheme = theme === 'dark' ? 'light' : 'dark';
    setTheme(nextTheme);
    localStorage.setItem('theme', nextTheme);
    if (nextTheme === 'dark') {
      document.documentElement.classList.add('dark');
      document.documentElement.classList.remove('light');
    } else {
      document.documentElement.classList.add('light');
      document.documentElement.classList.remove('dark');
    }
  };

  const t = T[lang] || T.en;
  const rtl = lang === 'ar';
  const isMgr = Boolean(profile?.is_manager || (profile?.role && profile.role !== 'MR'));
  const periodLabel = period;
  // Stable ref tracking what data has been fetched — survives re-renders without causing them
  const fetchedKeyRef = React.useRef(null);
  // Key currently being fetched. fetchedKeyRef is only set AFTER a load
  // finishes, so it cannot stop a second load that starts while the first
  // is still running -- which is how the dashboard ended up issuing two
  // full sets of paginated requests (102 instead of 51) on every open.
  const inFlightKeyRef = React.useRef(null);
  const codesKey = visibleCodes ? [...visibleCodes].sort().join(',') : '';
  const currentKey = `${periodLabel}|${codesKey}|${isMgr}`;

  const load = useCallback(async (force = false) => {
    if (!visibleCodes?.length) { setLoading(false); return; }
    const isAdmin = profile?.role === 'Admin';
    const codes = visibleCodes;
    // codesKey and isAdmin MUST be in the cache key. Without them a rep and
    // a manager in the same tab collide on one entry and see each other's data.
    const cacheKey = `dash_v3_${periodLabel}_${isMgr}_${isAdmin}_${codesKey}`;

    // Skip if already fetched this key (tab switch won't retrigger)
    if (!force && fetchedKeyRef.current === currentKey) return;
    // Already fetching this exact key -- do not start a second run.
    if (inFlightKeyRef.current === currentKey) return;
    inFlightKeyRef.current = currentKey;

    const SPECIAL_MANAGERS = [
      'ahmad morsy', 'ahmed elasyed', 'ahmed tarek mohamed', 'akram ahmed elhossary',
      'asmaa abdel fattah', 'dm', 'evette zakaria hefni', 'gihad sayed', 'hosney mohamed',
      'islam abd elrahman', 'kamel ragab', 'mahmoud essam', 'mahmoud rabee', 'mahmoud younis',
      'mohamed elmostafa', 'mohamed shenawey', 'reda hasan abdelmaksod', 'samr nabil',
      'ahmad behiery', 'tamer lamee', 'wael zaki'
    ];

    const overrideSpecialManagers = (rows) => (rows || []).map(r => {
      const name = (r.user_name || r.employee_name || r.manager_name || r.rep_name || '').toLowerCase();
      if (SPECIAL_MANAGERS.includes(name)) {
        return { ...r, team: 'Other Managers' };
      }
      return r;
    });

    // Try sessionStorage cache first
    const cached = !force && sessionStorage.getItem(cacheKey);
    if (cached) {
      try {
        const parsed = JSON.parse(cached);
        if (parsed.visits && parsed.visits.length > 3000) {
          setSummary(overrideSpecialManagers(parsed.summaries));
          setSpecialty(overrideSpecialManagers(parsed.specialty));
          setProducts(overrideSpecialManagers(parsed.products));
          setCoaching(overrideSpecialManagers(parsed.coaching));
          setVisits(parsed.visits);
          fetchedKeyRef.current = currentKey;
          inFlightKeyRef.current = null;
          setLoading(false);
          return;
        }
      } catch (e) { /* ignore corrupted cache */ }
    }

    setLoading(true); setError('');

    const parsePeriodToDates = (pStr) => {
      if (!pStr) return { startDate: null, endDate: null };
      const months = {
        january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
        july: 7, august: 8, september: 9, october: 10, november: 11, december: 12
      };
      const parts = pStr.trim().split(/\s+/);
      if (parts.length === 2) {
        const mName = parts[0].toLowerCase();
        const yr = parseInt(parts[1], 10);
        if (months[mName] && !isNaN(yr)) {
          const mNum = months[mName];
          const mo = String(mNum).padStart(2, '0');
          const startDate = `${yr}-${mo}-01`;
          const lastDay = new Date(yr, mNum, 0).getDate();
          const endDate = `${yr}-${mo}-${String(lastDay).padStart(2, '0')}`;
          return { startDate, endDate };
        }
      }
      return { startDate: null, endDate: null };
    };

    const fetchVisitsPaginated = async () => {
      const t0 = performance.now();
      const { startDate, endDate } = parsePeriodToDates(periodLabel);
      const pageSize = 1000;
      const selectCols = 'user,employee_code,visit_date,visit_time,shift,acc_type_category,acc_type_raw,visit_type_category,doctor_name,doctor_key,acc_name,acc_id,team,specialty,classification,products';

      // ORDER BY is mandatory. Postgres gives no stable row order without
      // one, so each .range() request is free to return rows in a different
      // sequence -- some rows arrive twice, others never. 'id' is the
      // tiebreaker: visit_date alone is not unique, and a non-unique sort
      // key is still unstable across requests.
      //
      // It is also what makes the parallel fetch below CORRECT. Firing all
      // page requests at once is only safe because every request now sorts
      // identically; without ORDER BY the pages would not line up.
      const buildQuery = () => {
        let q = supabase.from('visits').select(selectCols, { count: 'exact' });
        if (periodLabel) {
          q = q.eq('period', periodLabel);
        } else if (startDate && endDate) {
          q = q.gte('visit_date', startDate).lte('visit_date', endDate);
        }
        return q.order('visit_date', { ascending: true })
          .order('id', { ascending: true });
      };

      // One cheap request to learn how many rows exist, so the rest can go
      // out in parallel instead of six sequential round trips.
      const { count: total, error: countErr } = await buildQuery()
        .range(0, 0);
      if (countErr) throw new Error(`visits count failed: ${countErr.message}`);
      if (!total) return [];

      const pages = Math.min(Math.ceil(total / pageSize), 50);
      if (pages === 50) {
        console.warn('visits: hit the 50,000-row cap; data is truncated.');
      }

      const results = await Promise.all(
        Array.from({ length: pages }, (_, i) =>
          buildQuery()
            .range(i * pageSize, (i + 1) * pageSize - 1)
            .then(({ data, error }) => {
              if (error) throw new Error(`visits page ${i} failed: ${error.message}`);
              return data || [];
            })
        )
      );

      let allVisits = results.flat();

      if (allVisits.length < total) {
        console.error(
          `visits: fetched ${allVisits.length} of ${total} rows. ` +
          `Dashboard totals will be understated.`
        );
      }
      console.info(
        `visits: ${allVisits.length} rows in ${pages} parallel page(s), ` +
        `${Math.round(performance.now() - t0)}ms`
      );

      if (!isAdmin && codes && codes.length > 0) {
        const codeSet = new Set(codes.map(c => String(c).trim()));
        allVisits = allVisits.filter(v => v.employee_code && codeSet.has(String(v.employee_code).trim()));
      }
      return allVisits;
    };

    // try/finally is load-bearing. setLoading(true) happens above, and
    // fetchVisitsPaginated now throws on a failed page instead of quietly
    // returning short data. Without finally, one bad page leaves the
    // dashboard on 'Loading...' forever with nothing on screen.
    try {
      const [rpcRes, teamsRes, visitsData] = await Promise.all([
        supabase.rpc('get_dashboard_data', {
          p_period: periodLabel,
          p_codes: codes,
          p_is_admin: isAdmin,
          p_is_manager: isMgr
        }),
        supabase.from('teams').select('id, name'),
        fetchVisitsPaginated()
      ]);

      if (visitsData) setVisits(visitsData);

      if (teamsRes.data) {
        const tMap = {};
        teamsRes.data.forEach(t => tMap[t.id] = t.name);
        setTeamsMap(tMap);
      }

      const { data, error: rpcError } = rpcRes;

      if (rpcError) {
        setError(rpcError.message);
      } else {
        // The old `catch (e) { }` here hid a QuotaExceededError. The
        // visits array is ~2.5 MB for a SINGLE team-month, and
        // sessionStorage caps out around 5 MB, so for an Admin viewing
        // all teams the write always threw and was always swallowed --
        // the cache silently never persisted and every single load
        // refetched everything from scratch.
        data.visits = visitsData;
        try {
          sessionStorage.setItem(cacheKey, JSON.stringify(data));
        } catch (e) {
          // Drop the stale entry so a half-written value can't be read back.
          try { sessionStorage.removeItem(cacheKey); } catch (_) { }
          console.warn(
            `Dashboard cache disabled: payload too large for sessionStorage ` +
            `(${visitsData?.length ?? 0} visit rows). Every reload will refetch. ` +
            `Move the aggregation server-side to fix this properly.`
          );
        }
      }

      setSummary(overrideSpecialManagers(data?.summaries));
      setSpecialty(overrideSpecialManagers(data?.specialty));
      setProducts(overrideSpecialManagers(data?.products));
      setCoaching(overrideSpecialManagers(data?.coaching));
      fetchedKeyRef.current = currentKey;
    } catch (e) {
      // Surface it. A visible error beats an eternal spinner, and beats
      // silently-wrong totals even more.
      console.error('Dashboard load failed:', e);
      setError(e?.message || 'Failed to load dashboard data.');
    } finally {
      inFlightKeyRef.current = null;
      setLoading(false);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [periodLabel, currentKey, isMgr, profile]);

  useEffect(() => { load(); }, [load]);

  // ── Filtering ──────────────────────────────────────────────────────────────
  const [viewMode, setViewMode] = useState('teams'); // 'teams' | 'employees'

  const teams = useMemo(() => {
    const isAdmin = profile?.role === 'Admin';
    const set = new Set();
    summary.forEach(r => {
      const rawTms = (r.team || '').split('; ').filter(Boolean);
      if (!rawTms.length) {
        if (isAdmin) set.add(rtl ? 'مدراء آخرين' : 'Other Managers');
      } else {
        rawTms.forEach(t => {
          if (t === 'Unknown' || t === 'Other Managers') {
            if (isAdmin) set.add(rtl ? 'مدراء آخرين' : 'Other Managers');
          } else {
            set.add(t);
          }
        });
      }
    });
    return [...set].sort();
  }, [summary, profile, rtl]);

  const byTeam = useCallback(rows => {
    const isAdmin = profile?.role === 'Admin';
    if (team === 'all') {
      if (!isAdmin) {
        return rows.filter(r => r.team && r.team !== 'Unknown' && r.team !== 'Other Managers');
      }
      return rows;
    }
    if (team === 'Other Managers' || team === 'Unknown' || team === 'مدراء آخرين') {
      return rows.filter(r => !r.team || r.team === 'Unknown' || r.team === 'Other Managers' || r.team === 'مدراء آخرين');
    }
    return rows.filter(r => (r.team || '').split('; ').includes(team));
  }, [team, profile]);

  // userHierarchyMap previously defined here

  const allLineManagers = useMemo(() => {
    const list = new Set();
    (hierarchy || []).forEach(h => {
      if (h.blm_name && !h.blm_name.includes('Directory') && !h.blm_name.includes('TEAM')) {
        list.add(h.blm_name);
      }
    });
    return [...list].sort();
  }, [hierarchy]);

  const territoryEmployeeNamesMap = useMemo(() => {
    const map = {};
    if (!hierarchy?.length) return map;
    (hierarchy || []).forEach(h => {
      if (h.role === 'Area Manager' && h.division_name && h.employee_name) {
        const div = h.division_name.trim();
        const amName = h.employee_name;
        if (!map[div]) map[div] = new Set();
        map[div].add(amName);
        hierarchy.forEach(sub => {
          if (sub.area_manager_name === amName || sub.employee_name === amName) {
            if (sub.employee_name) map[div].add(sub.employee_name);
          }
        });
      }
    });
    return map;
  }, [hierarchy]);

  const allManagerTerritories = useMemo(() => {
    const list = new Set();
    const currentTeamUsers = new Set();
    if (team === 'all') {
      summary.forEach(r => currentTeamUsers.add(r.user_name));
    } else {
      summary.forEach(r => {
        if ((r.team || '').split('; ').includes(team)) {
          currentTeamUsers.add(r.user_name);
        }
      });
    }

    (hierarchy || []).forEach(h => {
      if (h.role === 'Area Manager' && h.division_name) {
        const divName = h.division_name.trim();
        if (team === 'all') {
          list.add(divName);
        } else {
          const namesInDiv = territoryEmployeeNamesMap[divName] || new Set();
          let hasOverlap = false;
          for (const n of namesInDiv) {
            if (currentTeamUsers.has(n)) {
              hasOverlap = true;
              break;
            }
          }
          if (hasOverlap) {
            list.add(divName);
          }
        }
      }
    });
    return [...list].sort();
  }, [hierarchy, team, summary, territoryEmployeeNamesMap]);

  const byLineManager = useCallback(rows => {
    if (lineManagerFilter === 'all') return rows;
    return rows.filter(r => {
      const name = r.user_name || r.employee_name || r.manager_name || r.rep_name;
      const userMeta = userHierarchyMap[(name || '').toLowerCase().trim()];
      return userMeta?.blm_name === lineManagerFilter;
    });
  }, [lineManagerFilter, userHierarchyMap]);

  const byManagerTerritory = useCallback(rows => {
    if (managerTerritoryFilter === 'all') return rows;
    const allowedNames = territoryEmployeeNamesMap[managerTerritoryFilter];
    return rows.filter(r => {
      const name = r.user_name || r.employee_name || r.manager_name || r.rep_name;
      if (allowedNames && allowedNames.has(name)) return true;
      if (r.territory && r.territory.includes(managerTerritoryFilter)) return true;
      if (r.division_name === managerTerritoryFilter) return true;
      return false;
    });
  }, [managerTerritoryFilter, territoryEmployeeNamesMap]);

  const fSummary = useMemo(() => {
    let r = byManagerTerritory(byLineManager(byTeam(summary)));
    if (search) r = r.filter(x => x.user_name?.toLowerCase().includes(search.toLowerCase()) || x.territory?.toLowerCase().includes(search.toLowerCase()));
    if (userFilter !== 'all') {
      const targetNames = new Set([userFilter]);
      (hierarchy || []).forEach(h => {
        if (h.area_manager_name === userFilter || h.supervisor_name === userFilter) {
          if (h.employee_name) targetNames.add(h.employee_name);
        }
        if (h.employee_name === userFilter) {
          if (h.supervisor_name) targetNames.add(h.supervisor_name);
          else if (h.area_manager_name) targetNames.add(h.area_manager_name);
          else if (h.blm_name && !h.blm_name.includes('Directory') && !h.blm_name.includes('TEAM')) targetNames.add(h.blm_name);
        }
      });
      r = r.filter(x => targetNames.has(x.user_name));
    }
    const m = new Map();
    r.forEach(x => {
      const k = x.user_name;
      if (!m.has(k)) {
        m.set(k, { ...x, _am_days_sum: x.am_shift_days || 0, _pm_days_sum: x.pm_shift_days || 0, _am_dur_sum: (x.avg_am_shift_hm || 0) * (x.am_shift_days || 0), _pm_dur_sum: (x.avg_pm_shift_hm || 0) * (x.pm_shift_days || 0) });
      } else {
        const existing = m.get(k);
        const sumKeys = ['working_days', 'complete_field_days', 'am_shift_days', 'pm_shift_days', 'double_visit_days', 'office_work_days', 'no_activities', 'no_events', 'am_calls', 'pm_calls', 'total_am_covered', 'total_pm_covered', 'amcenter_covered', 'hospital_covered', 'clinic_covered', 'polyclinic_covered', 'pharmacies_visited', 'pharmacies_covered', 'total_product_calls', 'distinct_products', 'coaching_days'];
        sumKeys.forEach(sk => existing[sk] = (existing[sk] || 0) + (x[sk] || 0));
        if (x.territory && !existing.territory?.includes(x.territory)) {
          existing.territory = existing.territory ? `${existing.territory}; ${x.territory}` : x.territory;
        }
        existing._am_days_sum += (x.am_shift_days || 0);
        existing._pm_days_sum += (x.pm_shift_days || 0);
        existing._am_dur_sum += (x.avg_am_shift_hm || 0) * (x.am_shift_days || 0);
        existing._pm_dur_sum += (x.avg_pm_shift_hm || 0) * (x.pm_shift_days || 0);
        existing.am_call_rate = existing.am_shift_days ? (existing.am_calls / existing.am_shift_days) : 0;
        existing.pm_call_rate = existing.pm_shift_days ? (existing.pm_calls / existing.pm_shift_days) : 0;
        existing.avg_am_shift_hm = existing._am_days_sum > 0 ? (existing._am_dur_sum / existing._am_days_sum) : 0;
        existing.avg_pm_shift_hm = existing._pm_days_sum > 0 ? (existing._pm_dur_sum / existing._pm_days_sum) : 0;
      }
    });
    const finalArr = Array.from(m.values());

    // Recalculate coaching days from actual coaching data (filtered by exact time grain date if set)
    const filteredCoaching = filterByTimeGrain(coaching || []);
    const mgrCoachingMap = {};
    const repCoachingMap = {};
    filteredCoaching.forEach(c => {
      const mgr = c.manager_name;
      const rep = c.rep_name;
      if (mgr) {
        if (!mgrCoachingMap[mgr]) mgrCoachingMap[mgr] = new Set();
        if (c.coaching_date) mgrCoachingMap[mgr].add(c.coaching_date);
      }
      if (rep) {
        if (!repCoachingMap[rep]) repCoachingMap[rep] = new Set();
        if (c.coaching_date) repCoachingMap[rep].add(c.coaching_date);
      }
    });

    finalArr.forEach(x => {
      if (x.is_manager) {
        x.coaching_days = mgrCoachingMap[x.user_name] ? mgrCoachingMap[x.user_name].size : 0;
      } else {
        x.coaching_days = repCoachingMap[x.user_name] ? repCoachingMap[x.user_name].size : 0;
      }
    });

    if (timeGrain !== 'all') {
      const hasVisitsData = (rawVisits || []).length > 0;
      const ratio = getTimeGrainRatio();
      const numKeys = ['working_days', 'complete_field_days', 'am_shift_days', 'pm_shift_days', 'double_visit_days', 'office_work_days', 'no_activities', 'no_events', 'am_calls', 'pm_calls', 'total_am_covered', 'total_pm_covered', 'amcenter_covered', 'hospital_covered', 'clinic_covered', 'polyclinic_covered', 'pharmacies_visited', 'pharmacies_covered', 'total_product_calls'];

      // Pre-filter and pre-group rawVisits once for O(1) representative lookup
      const visitsByCodeMap = {};
      const visitsByNameMap = {};
      if (hasVisitsData) {
        const slicedVisits = filterByTimeGrain(rawVisits);
        slicedVisits.forEach(v => {
          if (v.employee_code) {
            const c = String(v.employee_code).trim();
            if (!visitsByCodeMap[c]) visitsByCodeMap[c] = [];
            visitsByCodeMap[c].push(v);
          }
          if (v.user) {
            const u = v.user.toLowerCase().trim();
            if (!visitsByNameMap[u]) visitsByNameMap[u] = [];
            visitsByNameMap[u].push(v);
          }
        });
      }

      finalArr.forEach(x => {
        const normName = x.user_name?.toLowerCase().trim();
        const code = x.employee_code;
        const codeKey = code ? String(code).trim() : null;
        const userVisitsByCode = codeKey ? visitsByCodeMap[codeKey] : null;
        const userVisitsByName = normName ? visitsByNameMap[normName] : null;
        const userVisits = hasVisitsData ? (userVisitsByCode || userVisitsByName || []) : [];

        if (hasVisitsData) {
          // Identify activity/office work dates to exclude from working days
          const isActivity = v => {
            const cat1 = (v.acc_type_category || '').toLowerCase();
            const cat2 = (v.visit_type_category || '').toLowerCase();
            return cat1.includes('activity') || cat1.includes('office') || cat2.includes('activity') || cat2.includes('office');
          };
          const activityDates = new Set(userVisits.filter(isActivity).map(v => v.visit_date).filter(Boolean));

          if (userVisits.length > 0 || x.coaching_days > 0) {
            // Count calls only from real doctor/account visits
            const actualVisits = userVisits.filter(v => v.doctor_name || v.acc_name);
            const amVisits = actualVisits.filter(v => v.shift === 'AM');
            const pmVisits = actualVisits.filter(v => v.shift === 'PM');
            const amCalls = amVisits.length;
            const pmCalls = pmVisits.length;

            // For DAY counting: use ALL visits (any row = person was present that day)
            // This prevents holidays/blank rows from zeroing out a day that had real visits
            const allUserAmVisits = userVisits.filter(v => v.shift === 'AM');
            const allUserPmVisits = userVisits.filter(v => v.shift === 'PM');
            const amDates = new Set(allUserAmVisits.map(v => v.visit_date).filter(Boolean));
            const pmDates = new Set(allUserPmVisits.map(v => v.visit_date).filter(Boolean));
            const allDates = new Set(userVisits.map(v => v.visit_date).filter(Boolean));

            // Managers get shift days and working days from their Coaching Days
            if (x.is_manager) {
              const mgrCoaches = filteredCoaching.filter(c => c.manager_name === x.user_name);
              mgrCoaches.forEach(c => {
                if (c.coaching_date) {
                  allDates.add(c.coaching_date);
                  // If we have am/pm breakdown in coaching
                  if (c.am_visits > 0 || c.am_accompanied > 0) amDates.add(c.coaching_date);
                  if (c.pm_visits > 0 || c.pm_accompanied > 0) pmDates.add(c.coaching_date);
                  // Fallback: if no specific AM/PM is logged but it's a coaching day, assume both or at least AM
                  if (!c.am_visits && !c.am_accompanied && !c.pm_visits && !c.pm_accompanied) {
                    amDates.add(c.coaching_date);
                  }
                }
              });
            }

            // Remove activity days from ALL day calculations (per user request)
            activityDates.forEach(d => {
              allDates.delete(d);
              amDates.delete(d);
              pmDates.delete(d);
            });

            let completeCount = 0;
            allDates.forEach(d => { if (amDates.has(d) && pmDates.has(d)) completeCount++; });

            x.am_calls = amCalls;
            x.pm_calls = pmCalls;
            x.am_shift_days = amDates.size;
            x.pm_shift_days = pmDates.size;
            x.working_days = allDates.size;
            x.complete_field_days = completeCount;
            // Zero out fields that we can't accurately slice from rawVisits alone
            x.office_work_days = 0;
            x.no_activities = 0;
            x.no_events = 0;
            x.double_visit_days = 0;
            x.avg_am_shift_hm = 0;
            x.avg_pm_shift_hm = 0;
            x.avg_field_overall_hm = 0;
            x.total_visits = userVisits.length;
            x.am_call_rate = amDates.size > 0 ? Math.round((amCalls / amDates.size) * 10) / 10 : 0;
            x.pm_call_rate = pmDates.size > 0 ? Math.round((pmCalls / pmDates.size) * 10) / 10 : 0;
            x.total_am_covered = new Set(amVisits.map(v => v.doctor_key || v.doctor_name).filter(Boolean)).size;
            x.total_pm_covered = new Set(pmVisits.map(v => v.doctor_key || v.doctor_name).filter(Boolean)).size;
            // AM account coverage: unique accounts (acc_id or acc_name) for Hospital/AM Center
            const amAccountVisits = amVisits.filter(v => {
              const cat = (v.acc_type_category || '').toLowerCase();
              return cat.includes('hospital') || cat.includes('am center');
            });
            const amAccountIds = amAccountVisits.map(v => v.acc_id || v.acc_name).filter(Boolean);
            x.am_accounts_unique = new Set(amAccountIds).size;
            x.am_accounts_revisits = Math.max(0, amAccountIds.length - new Set(amAccountIds).size);
            x.pharmacies_visited = userVisits.filter(v =>
              (v.acc_type_category || '').toLowerCase().includes('pharmacy') ||
              (v.acc_type_raw || '').toLowerCase().includes('pharmacy') ||
              (v.acc_name || '').toLowerCase().includes('pharmacy')
            ).length;
            x.amcenter_covered = new Set(amVisits.filter(v => (v.acc_type_category || '').toLowerCase().includes('am center')).map(v => v.doctor_key || v.doctor_name).filter(Boolean)).size;
            x.hospital_covered = new Set(amVisits.filter(v => (v.acc_type_category || '').toLowerCase().includes('hospital')).map(v => v.doctor_key || v.doctor_name).filter(Boolean)).size;
            x.clinic_covered = new Set(pmVisits.filter(v => {
              const cat = (v.acc_type_category || '').toLowerCase();
              return cat.includes('clinic') && !cat.includes('poly');
            }).map(v => v.doctor_key || v.doctor_name).filter(Boolean)).size;
            x.polyclinic_covered = new Set(pmVisits.filter(v => (v.acc_type_category || '').toLowerCase().includes('poly')).map(v => v.doctor_key || v.doctor_name).filter(Boolean)).size;
            x.pharmacies_covered = new Set(userVisits.filter(v => {
              const cat = (v.acc_type_category || '').toLowerCase();
              const raw = (v.acc_type_raw || '').toLowerCase();
              const name = (v.acc_name || '').toLowerCase();
              return cat.includes('pharmacy') || raw.includes('pharmacy') || name.includes('pharmacy');
            }).map(v => v.acc_id || v.acc_name).filter(Boolean)).size;
            x.total_product_calls = 0;
            x.distinct_products = 0;
          } else {
            x.am_calls = 0;
            x.pm_calls = 0;
            x.am_shift_days = 0;
            x.pm_shift_days = 0;
            x.working_days = 0;
            x.complete_field_days = 0;
            x.office_work_days = 0;
            x.no_activities = 0;
            x.no_events = 0;
            x.double_visit_days = 0;
            x.avg_am_shift_hm = 0;
            x.avg_pm_shift_hm = 0;
            x.avg_field_overall_hm = 0;
            x.total_visits = 0;
            x.am_call_rate = 0;
            x.pm_call_rate = 0;
            x.total_am_covered = 0;
            x.total_pm_covered = 0;
            x.am_accounts_unique = 0;
            x.am_accounts_revisits = 0;
            x.pharmacies_visited = 0;
            x.amcenter_covered = 0;
            x.hospital_covered = 0;
            x.clinic_covered = 0;
            x.polyclinic_covered = 0;
            x.pharmacies_covered = 0;
            x.total_product_calls = 0;
            x.distinct_products = 0;
          }
        } else {
          // Fallback if visits table isn't populated yet in database
          numKeys.forEach(sk => {
            if (x[sk]) x[sk] = Math.max(1, Math.round(x[sk] * ratio));
          });
          x.am_call_rate = x.am_shift_days ? Math.round((x.am_calls / x.am_shift_days) * 10) / 10 : 0;
          x.pm_call_rate = x.pm_shift_days ? Math.round((x.pm_calls / x.pm_shift_days) * 10) / 10 : 0;
        }
      });
    }

    return sortSummary(finalArr);
  }, [summary, coaching, rawVisits, byTeam, byLineManager, byManagerTerritory, search, userFilter, hierarchy, timeGrain, filterByTimeGrain]);

  const managerNames = useMemo(() => {
    const s = new Set();
    summary.forEach(r => {
      if (r.is_manager) s.add((r.user_name || '').toLowerCase().trim());
    });
    return s;
  }, [summary]);

  const fSpecialty = useMemo(() => {
    let sourceData = specialty;
    if (timeGrain !== 'all' && rawVisits?.length > 0) {
      const filteredV = filterByTimeGrain(rawVisits);
      const groups = {};
      filteredV.forEach(v => {
        if (!v.shift || (v.shift !== 'AM' && v.shift !== 'PM')) return;
        const uName = v.user || '';
        const spec = v.specialty || 'Unknown';
        const cls = v.classification || 'Unknown';
        const key = `${uName}||${spec}||${cls}||${v.shift}`;
        if (!groups[key]) {
          groups[key] = {
            user_name: uName,
            employee_code: v.employee_code,
            specialty: spec,
            classification: cls,
            shift: v.shift,
            team: v.team || '',
            call_count: 0,
            _docs: new Set()
          };
        }
        groups[key].call_count += 1;
        if (v.doctor_key || v.doctor_name) groups[key]._docs.add(v.doctor_key || v.doctor_name);
      });
      sourceData = Object.values(groups).map(g => ({
        ...g,
        unique_doctors: g._docs.size
      }));
    }

    let r = byManagerTerritory(byLineManager(byTeam(sourceData)));
    if (search) r = r.filter(x => x.user_name?.toLowerCase().includes(search.toLowerCase()) || x.territory?.toLowerCase().includes(search.toLowerCase()));
    if (userFilter !== 'all') {
      const targetNames = new Set([userFilter]);
      (hierarchy || []).forEach(h => {
        if (h.area_manager_name === userFilter || h.supervisor_name === userFilter) {
          if (h.employee_name) targetNames.add(h.employee_name);
        }
        if (h.employee_name === userFilter) {
          if (h.supervisor_name) targetNames.add(h.supervisor_name);
          else if (h.area_manager_name) targetNames.add(h.area_manager_name);
          else if (h.blm_name && !h.blm_name.includes('Directory') && !h.blm_name.includes('TEAM')) targetNames.add(h.blm_name);
        }
      });
      r = r.filter(x => targetNames.has(x.user_name));
    }

    // Exclude managers from team totals unless explicitly filtered to a single manager
    if (userFilter === 'all') {
      r = r.filter(x => !managerNames.has((x.user_name || '').toLowerCase().trim()));
    }

    return r;
  }, [specialty, rawVisits, timeGrain, filterByTimeGrain, byTeam, byLineManager, byManagerTerritory, search, userFilter, hierarchy, managerNames]);

  const fProducts = useMemo(() => {
    let sourceData = products;
    if (timeGrain !== 'all' && rawVisits?.length > 0) {
      const filteredV = filterByTimeGrain(rawVisits);
      const groups = {};
      filteredV.forEach(v => {
        if (!v.shift || (v.shift !== 'AM' && v.shift !== 'PM') || !v.products) return;
        const uName = v.user || '';
        const spec = v.specialty || 'Unknown';
        const prods = String(v.products).split(',');
        prods.forEach(pRaw => {
          const prod = pRaw.trim();
          if (!prod) return;
          const key = `${uName}||${prod}||${v.shift}||${spec}`;
          if (!groups[key]) {
            groups[key] = {
              user_name: uName,
              employee_code: v.employee_code,
              product: prod,
              shift: v.shift,
              specialty: spec,
              team: v.team || '',
              call_count: 0,
              _docs: new Set()
            };
          }
          groups[key].call_count += 1;
          if (v.doctor_key || v.doctor_name) groups[key]._docs.add(v.doctor_key || v.doctor_name);
        });
      });
      sourceData = Object.values(groups).map(g => ({
        ...g,
        unique_doctors: g._docs.size
      }));
    }

    let r = byManagerTerritory(byLineManager(byTeam(sourceData)));
    if (search) r = r.filter(x => x.user_name?.toLowerCase().includes(search.toLowerCase()) || x.territory?.toLowerCase().includes(search.toLowerCase()));
    if (userFilter !== 'all') {
      const targetNames = new Set([userFilter]);
      (hierarchy || []).forEach(h => {
        if (h.area_manager_name === userFilter || h.supervisor_name === userFilter) {
          if (h.employee_name) targetNames.add(h.employee_name);
        }
        if (h.employee_name === userFilter) {
          if (h.supervisor_name) targetNames.add(h.supervisor_name);
          else if (h.area_manager_name) targetNames.add(h.area_manager_name);
          else if (h.blm_name && !h.blm_name.includes('Directory') && !h.blm_name.includes('TEAM')) targetNames.add(h.blm_name);
        }
      });
      r = r.filter(x => targetNames.has(x.user_name));
    }

    // Exclude managers from team totals unless explicitly filtered to a single manager
    if (userFilter === 'all') {
      r = r.filter(x => !managerNames.has((x.user_name || '').toLowerCase().trim()));
    }

    return r;
  }, [products, rawVisits, timeGrain, filterByTimeGrain, byTeam, byLineManager, byManagerTerritory, search, userFilter, hierarchy, managerNames]);

  const visibleNames = useMemo(() => {
    if (!hierarchy?.length || !visibleCodes?.length) return null;
    const vCodesSet = new Set(visibleCodes);
    const names = new Set();
    if (profile?.employee_name) names.add(profile.employee_name);
    hierarchy.forEach(h => {
      if (vCodesSet.has(h.employee_code) && h.employee_name) {
        names.add(h.employee_name);
      }
    });
    return names;
  }, [hierarchy, visibleCodes, profile]);

  const fCoaching = useMemo(() => {
    let r = byManagerTerritory(byLineManager(byTeam(filterByTimeGrain(coaching))));
    if (visibleNames && profile?.role !== 'Admin') {
      r = r.filter(x => visibleNames.has(x.manager_name) || visibleNames.has(x.rep_name));
    }
    if (search) r = r.filter(x => x.manager_name?.toLowerCase().includes(search.toLowerCase()) || x.rep_name?.toLowerCase().includes(search.toLowerCase()));
    if (userFilter !== 'all') {
      const managerNames = new Set([userFilter]);
      (hierarchy || []).forEach(h => {
        if (h.area_manager_name === userFilter || h.supervisor_name === userFilter) {
          if (h.role === 'Area Manager' || h.role === 'Supervisor' || h.employee_name === userFilter) {
            if (h.employee_name) managerNames.add(h.employee_name);
          }
        }
        if (h.employee_name === userFilter) {
          if (h.supervisor_name) managerNames.add(h.supervisor_name);
          else if (h.area_manager_name) managerNames.add(h.area_manager_name);
          else if (h.blm_name && !h.blm_name.includes('Directory') && !h.blm_name.includes('TEAM')) managerNames.add(h.blm_name);
        }
      });
      r = r.filter(x => managerNames.has(x.manager_name) || managerNames.has(x.rep_name));
    }
    return r;
  }, [coaching, filterByTimeGrain, byTeam, byLineManager, byManagerTerritory, search, userFilter, visibleNames, profile, hierarchy]);

  // ── Timing data: last visit time per rep per day (PM ONLY) ──────────────────────
  const timingData = useMemo(() => {
    if (!rawVisits?.length) return [];

    // Identify dates where a user had a PM Activity or PM Office Work
    const pmActivityDates = new Set();
    rawVisits.forEach(v => {
      if (v.shift === 'PM') {
        const cat1 = (v.acc_type_category || '').toLowerCase();
        const cat2 = (v.visit_type_category || '').toLowerCase();
        const raw = (v.acc_type_raw || '').toLowerCase();
        const isActivityOrOffice = cat1.includes('activity') || cat1.includes('office') ||
          cat2.includes('activity') || cat2.includes('office') ||
          raw.includes('activity') || raw.includes('office');
        if (isActivityOrOffice && v.user && v.visit_date) {
          pmActivityDates.add(`${v.user}|||${v.visit_date}`);
        }
      }
    });

    // Condition: Last recorded PM visit ONLY (shift === 'PM'), ignoring AM visits
    const validVisits = rawVisits.filter(v => v.visit_date && v.visit_time && v.shift === 'PM');

    // Group by user + date, find latest time
    const byUserDate = {};
    validVisits.forEach(v => {
      const user = v.user || '';
      const date = v.visit_date || '';
      if (!user || !date) return;
      const key = `${user}|||${date}`;

      // Exclude this date entirely for this user if they had PM activity or PM office work that day!
      if (pmActivityDates.has(key)) return;

      const time = v.visit_time || '';
      if (!byUserDate[key] || time > byUserDate[key].time) {
        byUserDate[key] = { user, date, time, team: v.team || '', employee_code: v.employee_code };
      }
    });

    // Categorize each entry
    const categorizeTime = (timeStr) => {
      if (!timeStr) return 'unknown';
      const mins = parseTimeToMinutes(timeStr);
      if (mins === null) return 'unknown';
      if (mins < 15 * 60) return 'early';    // before 3:00 PM (15:00)
      if (mins <= 18 * 60) return 'normal';   // 3:00 PM - 6:00 PM (15:00-18:00)
      return 'late';                           // after 6:00 PM (18:00)
    };

    return Object.values(byUserDate).map(entry => ({
      ...entry,
      category: categorizeTime(entry.time),
      formattedTime: entry.time ? formatMinutesToTime(parseTimeToMinutes(entry.time)) : '—',
    }));
  }, [rawVisits]);

  // Filtered timing data (same team/search/user filters as other tabs)
  const fTiming = useMemo(() => {
    let r = timingData;
    // Team filter
    if (team !== 'all') {
      if (team === 'Other Managers' || team === 'Unknown' || team === 'مدراء آخرين') {
        r = r.filter(x => !x.team || x.team === 'Unknown' || x.team === 'Other Managers');
      } else {
        r = r.filter(x => (x.team || '').split('; ').includes(team));
      }
    }
    // Enrich team from userTeamMap
    r = r.map(x => ({ ...x, team: userTeamMap[x.user?.toLowerCase().trim()] || x.team }));
    // Search filter
    if (search) r = r.filter(x => x.user?.toLowerCase().includes(search.toLowerCase()));
    // User filter
    if (userFilter !== 'all') r = r.filter(x => x.user === userFilter);
    // Time grain filter
    if (timeGrain !== 'all') r = filterByTimeGrain(r.map(x => ({ ...x, visit_date: x.date }))).map(x => ({ ...x }));
    // Line manager filter
    if (lineManagerFilter !== 'all') {
      r = r.filter(x => {
        const userMeta = userHierarchyMap[(x.user || '').toLowerCase().trim()];
        return userMeta?.blm_name === lineManagerFilter;
      });
    }
    // Manager territory filter
    if (managerTerritoryFilter !== 'all') {
      const allowedNames = territoryEmployeeNamesMap[managerTerritoryFilter];
      r = r.filter(x => allowedNames?.has(x.user));
    }
    return r.sort((a, b) => (a.date || '').localeCompare(b.date || '') || (a.user || '').localeCompare(b.user || ''));
  }, [timingData, team, search, userFilter, timeGrain, filterByTimeGrain, userTeamMap, lineManagerFilter, userHierarchyMap, managerTerritoryFilter, territoryEmployeeNamesMap]);

  // Timing summary stats
  const timingStats = useMemo(() => {
    const total = fTiming.length;
    const early = fTiming.filter(r => r.category === 'early').length;
    const normal = fTiming.filter(r => r.category === 'normal').length;
    const late = fTiming.filter(r => r.category === 'late').length;
    const uniqueDays = new Set(fTiming.map(r => r.date)).size;
    return { total, early, normal, late, uniqueDays };
  }, [fTiming]);

  // Timing category filter state & display limit for ultra-fast rendering
  const [timingCategoryFilter, setTimingCategoryFilter] = useState('all');
  const [timingDisplayLimit, setTimingDisplayLimit] = useState(100);

  const filteredTiming = useMemo(() => {
    if (timingCategoryFilter === 'all') return fTiming;
    return fTiming.filter(r => r.category === timingCategoryFilter);
  }, [fTiming, timingCategoryFilter]);

  const companyAverages = useMemo(() => {
    const reps = summary.filter(r => !r.is_manager);
    const avgs = {};
    NUMERIC_KPI_KEYS.forEach(key => {
      const vals = reps.map(r => Number(r[key]) || 0).filter(v => v > 0);
      avgs[key] = vals.length ? (vals.reduce((s, v) => s + v, 0) / vals.length) : 0;
    });
    return avgs;
  }, [summary]);

  // userTeamMap previously defined here

  const pmCoveragePieData = useMemo(() => {
    let rows = fSummary.filter(r => !r.is_manager);
    if (selectedRep) {
      rows = rows.filter(r => r.user_name === selectedRep);
    }
    const totalClinic = rows.reduce((s, r) => s + (Number(r.clinic_covered) || 0), 0);
    const totalPolyClinic = rows.reduce((s, r) => s + (Number(r.polyclinic_covered) || 0), 0);
    if (!totalClinic && !totalPolyClinic) return [];
    return [
      { label: rtl ? 'Clinic (عيادات)' : 'Clinic', value: totalClinic, color: '#1a6fc4' },
      { label: rtl ? 'Poly Clinic (مراكز)' : 'Poly Clinic', value: totalPolyClinic, color: '#10b981' }
    ];
  }, [fSummary, selectedRep, rtl]);

  const activityEventData = useMemo(() => {
    let rows = fSummary.filter(r => !r.is_manager);
    if (selectedRep) rows = rows.filter(r => r.user_name === selectedRep);
    const totalActivities = rows.reduce((s, r) => s + (Number(r.no_activities) || 0), 0);
    const totalEvents = rows.reduce((s, r) => s + (Number(r.no_events) || 0), 0);
    return { totalActivities, totalEvents };
  }, [fSummary, selectedRep]);

  const allUsers = useMemo(() => [...new Set(byTeam(summary).map(r => r.user_name))].sort(), [summary, byTeam]);
  const teamCount = new Set(fSummary.map(r => r.team)).size;

  const teamGroups = useMemo(() => {
    const isAdmin = profile?.role === 'Admin';
    const otherLabel = rtl ? 'مدراء آخرين' : 'Other Managers';
    if (team !== 'all') {
      const displayLabel = (team === 'Unknown' || team === 'Other Managers' || team === 'مدراء آخرين') ? otherLabel : (team || 'Team');
      return [{ label: displayLabel, rows: fSummary }];
    }
    const groups = {};
    fSummary.forEach(r => {
      const rawTms = (r.team && r.team !== 'Unknown') ? r.team.split('; ') : ['Unknown'];
      rawTms.forEach(tm => {
        let label = tm;
        if (tm === 'Unknown' || tm === 'Other Managers' || tm === 'مدراء آخرين') {
          if (!isAdmin) return; // Hide from non-admins
          label = otherLabel;
        }
        if (!groups[label]) groups[label] = [];
        groups[label].push(r);
      });
    });
    return Object.entries(groups).sort((a, b) => a[0].localeCompare(b[0])).map(([label, rows]) => ({ label, rows }));
  }, [fSummary, team, profile, rtl]);

  const visibleTabs = useMemo(() => {
    const all = Object.entries(t.tabs);
    return isMgr ? all : all.filter(([k]) => k !== 'coaching');
  }, [t.tabs, isMgr]);

  // ── Sidebar computed data ──────────────────────────────────────────────────
  const selectedRepData = useMemo(() => {
    if (!selectedRep) return null;
    return fSummary.find(r => r.user_name === selectedRep) || null;
  }, [fSummary, selectedRep]);

  // Specialty pie charts (shift-filtered & rep-filtered)
  const shiftFilteredSpecialty = useMemo(() => {
    let list = fSpecialty;
    if (selectedRep) {
      list = list.filter(r => r.user_name === selectedRep);
    }
    return shift === 'all' ? list : list.filter(r => r.shift === shift);
  }, [fSpecialty, shift, selectedRep]);

  const specialtyPieData = useMemo(() => {
    let list = shiftFilteredSpecialty;
    if (classificationFilter.size > 0) {
      list = list.filter(r => classificationFilter.has(r.classification));
    }
    const m = {};
    list.forEach(r => { const s = r.specialty || 'Other'; m[s] = (m[s] || 0) + (r.call_count || 0); });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, 10)
      .map(([label, value], i) => ({ label, value, color: PIE_COLORS[i % PIE_COLORS.length] }));
  }, [shiftFilteredSpecialty, classificationFilter]);

  const classificationPieData = useMemo(() => {
    let list = shiftFilteredSpecialty;
    if (specialtyFilter.size > 0) {
      list = list.filter(r => specialtyFilter.has(r.specialty));
    }
    const m = {};
    list.forEach(r => { const c = r.classification || 'Unclassified'; m[c] = (m[c] || 0) + (r.call_count || 0); });
    return Object.entries(m).sort((a, b) => b[1] - a[1])
      .map(([label, value], i) => ({ label, value, color: PIE_COLORS[i % PIE_COLORS.length] }));
  }, [shiftFilteredSpecialty, specialtyFilter]);

  const allSpecialties = useMemo(() =>
    [...new Set(fSpecialty.map(r => r.specialty).filter(Boolean))].sort()
    , [fSpecialty]);

  const filteredSpecialty = useMemo(() => {
    let res = fSpecialty;
    if (specialtyFilter.size > 0) {
      res = res.filter(r => specialtyFilter.has(r.specialty));
    }
    if (classificationFilter.size > 0) {
      res = res.filter(r => classificationFilter.has(r.classification));
    }
    return res;
  }, [fSpecialty, specialtyFilter, classificationFilter]);

  const filteredProducts = useMemo(() => {
    if (productFilter.size === 0) return fProducts;
    return fProducts.filter(r => productFilter.has(r.product));
  }, [fProducts, productFilter]);

  const allProducts = useMemo(() =>
    [...new Set(fProducts.map(r => r.product).filter(Boolean))].sort()
    , [fProducts]);

  // Products pie chart (shift-filtered & rep-filtered)
  const shiftFilteredProducts = useMemo(() => {
    let list = fProducts;
    if (selectedRep) {
      list = list.filter(r => r.user_name === selectedRep);
    }
    return shift === 'all' ? list : list.filter(r => r.shift === shift);
  }, [fProducts, shift, selectedRep]);

  const productPieData = useMemo(() => {
    const m = {};
    shiftFilteredProducts.forEach(r => { const p = r.product || 'Other'; m[p] = (m[p] || 0) + (r.call_count || 0); });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, 10)
      .map(([label, value], i) => ({ label, value, color: PIE_COLORS[i % PIE_COLORS.length] }));
  }, [shiftFilteredProducts]);

  const topProducts = useMemo(() => {
    const m = {};
    shiftFilteredProducts.forEach(r => { const p = r.product || 'Other'; m[p] = (m[p] || 0) + (r.call_count || 0); });
    const sorted = Object.entries(m).sort((a, b) => b[1] - a[1]);
    const max = sorted[0]?.[1] || 1;
    return sorted.slice(0, 8).map(([name, count]) => ({ name, count, pct: Math.round(count / max * 100) }));
  }, [shiftFilteredProducts]);

  // Coaching manager groups
  const managerGroups = useMemo(() => {
    const m = {};
    fCoaching.forEach(r => {
      const mgr = r.manager_name || 'Unknown';
      if (!m[mgr]) m[mgr] = { name: mgr, team: r.team || '', dates: new Set(), reps: new Set() };
      m[mgr].dates.add(r.coaching_date);
      m[mgr].reps.add(r.rep_name);
    });
    return Object.values(m).sort((a, b) => a.team.localeCompare(b.team) || a.name.localeCompare(b.name))
      .map(g => ({ ...g, dayCount: g.dates.size, repCount: g.reps.size }));
  }, [fCoaching]);

  const filteredCoaching = useMemo(() => {
    if (!selectedManager) return fCoaching;
    return fCoaching.filter(r => r.manager_name === selectedManager || r.rep_name === selectedManager || (visibleNames?.has(r.manager_name) && profile?.employee_name === selectedManager));
  }, [fCoaching, selectedManager, visibleNames, profile]);

  // ── Actions ────────────────────────────────────────────────────────────────
  function toggleProduct(p) {
    setProductFilter(prev => {
      const next = new Set(prev);
      if (next.has(p)) next.delete(p); else next.add(p);
      return next;
    });
  }

  const handleSelectRep = (repName) => {
    if (selectedRep === repName) {
      setSelectedRep(null);
      setUser('all');
    } else {
      setSelectedRep(repName);
      setUser(repName);
    }
  };

  function changeTab(k) {
    setTab(k);
    setSidebarOpen(false);
  }

  function doExport() {
    const wb = XLSX.utils.book_new();

    if (tab === 'timing') {
      const sh = [['Date', 'User Name', 'Team', 'Last Visit', 'Category']];
      [...filteredTiming].forEach(r => {
        sh.push([
          r.date || '—',
          r.user || '—',
          r.team || '—',
          r.formattedTime || '—',
          r.category === 'early' ? (t.kpi.timing_early || '< 3 PM')
            : r.category === 'normal' ? (t.kpi.timing_normal || '3–6 PM')
              : r.category === 'late' ? (t.kpi.timing_late || '> 6 PM')
                : '—'
        ]);
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sh), 'Last Visit Data');
      XLSX.writeFile(wb, `excellence_last_visit_data_${periodLabel.replace(' ', '_')}_${Date.now()}.xlsx`);
      return;
    }

    if (tab === 'coaching') {
      const rows = selectedManager ? filteredCoaching : fCoaching;
      const sh = [['Manager', 'Rep', 'Date', 'Team', 'AM Visits', 'AM Acc.', 'AM %', 'PM Visits', 'PM Acc.', 'PM %']];
      [...rows]
        .sort((a, b) => (a.manager_name || '').localeCompare(b.manager_name || '') || (a.coaching_date || '').localeCompare(b.coaching_date || ''))
        .forEach(r => sh.push([
          r.manager_name || '—', r.rep_name || '—', r.coaching_date || '—', r.team || '—',
          r.am_visits || 0, r.am_accompanied || 0, r.am_visits ? Math.round((r.am_accompanied / r.am_visits) * 100) + '%' : '-',
          r.pm_visits || 0, r.pm_accompanied || 0, r.pm_visits ? Math.round((r.pm_accompanied / r.pm_visits) * 100) + '%' : '-'
        ]));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sh), 'Coaching');
      XLSX.writeFile(wb, `excellence_coaching_${periodLabel.replace(' ', '_')}_${Date.now()}.xlsx`);
      return;
    }

    if (tab === 'specialty') {
      const sh = [['Specialty', 'User Name', 'Team', 'Territory', 'AM Calls', 'PM Calls', 'Total Calls']];
      const specMap = new Map();
      fSpecialty.forEach(r => {
        const specName = r.specialty || r.name || '—';
        const key = `${specName}||${r.user_name || '—'}`;
        if (!specMap.has(key)) {
          specMap.set(key, {
            specialty: specName,
            user_name: r.user_name || '—',
            team: r.team || '—',
            territory: r.territory || '—',
            am_calls: 0,
            pm_calls: 0
          });
        }
        const item = specMap.get(key);
        const count = Number(r.call_count || r.calls || r.v || 0);
        if (r.shift === 'AM') item.am_calls += count;
        else if (r.shift === 'PM') item.pm_calls += count;
        else item.am_calls += count;
      });

      Array.from(specMap.values()).forEach(item => {
        sh.push([
          item.specialty,
          item.user_name,
          item.team,
          item.territory,
          item.am_calls,
          item.pm_calls,
          item.am_calls + item.pm_calls
        ]);
      });

      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sh), 'Specialty');
      XLSX.writeFile(wb, `excellence_specialty_${periodLabel.replace(' ', '_')}_${Date.now()}.xlsx`);
      return;
    }

    if (tab === 'products') {
      const sh = [['Product Name', 'User Name', 'Team', 'Territory', 'AM Calls', 'PM Calls', 'Total Calls']];
      const prodMap = new Map();
      fProducts.forEach(r => {
        const prodName = r.product || r.product_name || r.name || '—';
        const key = `${prodName}||${r.user_name || '—'}`;
        if (!prodMap.has(key)) {
          prodMap.set(key, {
            product: prodName,
            user_name: r.user_name || '—',
            team: r.team || '—',
            territory: r.territory || '—',
            am_calls: 0,
            pm_calls: 0
          });
        }
        const item = prodMap.get(key);
        const count = Number(r.call_count || r.calls || r.v || 0);
        if (r.shift === 'AM') item.am_calls += count;
        else if (r.shift === 'PM') item.pm_calls += count;
        else item.am_calls += count;
      });

      Array.from(prodMap.values()).forEach(item => {
        sh.push([
          item.product,
          item.user_name,
          item.team,
          item.territory,
          item.am_calls,
          item.pm_calls,
          item.am_calls + item.pm_calls
        ]);
      });

      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sh), 'Products');
      XLSX.writeFile(wb, `excellence_products_${periodLabel.replace(' ', '_')}_${Date.now()}.xlsx`);
      return;
    }

    // Default: Summary tab (also always includes Team Averages)
    const allKpiKeys = t.kpiGroups.flatMap(g => g.keys);
    const sh = [['Team', 'User', 'Territory', 'Role', ...allKpiKeys.map(k => t.kpi[k] || k)]];
    fSummary.forEach(r => sh.push([r.team || '—', r.user_name || '—', r.territory || '—', r.role || (r.is_manager ? 'Manager' : 'MR'), ...allKpiKeys.map(k => r[k] ?? '')]));
    const aggRows = [['Team', 'KPI', 'Sum', 'Avg']];
    teamGroups.forEach(({ label, rows }) => {
      const { agg } = computeAggregates(rows);
      NUMERIC_KPI_KEYS.forEach(k => {
        if (agg[k]) aggRows.push([label, t.kpi[k] || k, agg[k].sum, +agg[k].avg.toFixed(2)]);
      });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sh), 'Summary');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aggRows), 'Team Averages');
    XLSX.writeFile(wb, `excellence_summary_${periodLabel.replace(' ', '_')}_${Date.now()}.xlsx`);
  }

  // ── Render ─────────────────────────────────────────────────────────────────
  const handleAiSubmit = async (e) => {
    e.preventDefault();
    if (!aiInput.trim() || isAiLoading) return;
    const msg = aiInput.trim();
    setAiInput('');
    setAiHistory(prev => [...prev, { role: 'user', content: msg }]);
    setIsAiLoading(true);

    try {
      const contextData = {
        tab,
        summary: fSummary,
        specialty: tab === 'specialty' ? fSpecialty : undefined,
        products: tab === 'products' ? fProducts : undefined,
      };

      const res = await fetch('/api/ai', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          messages: [...aiHistory, { role: 'user', content: msg }],
          contextData
        })
      });
      const data = await res.json();
      if (data.reply) {
        setAiHistory(prev => [...prev, { role: 'assistant', content: data.reply }]);
      } else {
        setAiHistory(prev => [...prev, { role: 'assistant', content: 'Sorry, I encountered an error.' }]);
      }
    } catch (err) {
      setAiHistory(prev => [...prev, { role: 'assistant', content: 'Failed to connect to AI.' }]);
    }
    setIsAiLoading(false);
  };

  return (
    <div className={`dash${rtl ? ' rtl' : ''}`} dir={rtl ? 'rtl' : 'ltr'}>

      {/* HEADER */}
      <header className="dash-hdr">
        <div className="dash-hdr-l">
          <div className="dash-brand-wrap">
            <span className="dash-brand">{rtl ? 'إكسيلنس - CRM' : 'Excellence - CRM'}</span>
            <span className="dash-brand-sub">{rtl ? 'تطبيق الويب' : 'web app'}</span>
          </div>
          <div className="dash-sep" />
          <span className="dash-view">{t.roleView[profile?.role] || ''}</span>
        </div>
        <div className="dash-hdr-r">
          {profile?.role === 'Admin' && <a className="hbtn hbtn-outline" href="#/admin">{t.adminPanel}</a>}
          <button className="hbtn hbtn-outline" style={{ padding: '6px 12px', fontSize: '13px' }} onClick={toggleTheme} title="Toggle Dark/Light Mode">
            {theme === 'dark' ? '☀️ Light' : '🌙 Dark'}
          </button>
          <button className="hbtn hbtn-lang" onClick={() => setLang(lang === 'en' ? 'ar' : 'en')}>
            {lang === 'en' ? 'عربي' : 'EN'}
          </button>
          <div className="dash-user">
            <div className="du-name">{profile?.employee_name}</div>
            <div className="du-role">{profile?.role} · {profile?.employee_code}</div>
          </div>
          <button
            className="hbtn hbtn-outline"
            title={rtl ? 'تحديث البيانات' : 'Refresh Data'}
            style={{ padding: '6px 10px', fontSize: '13px', lineHeight: 1 }}
            onClick={() => {
              sessionStorage.clear();
              fetchedKeyRef.current = null;
              setTeam('all');
              setShift('all');
              setSearch('');
              setUser('all');
              setSpecialtyFilter(new Set());
              setProductFilter(new Set());
              setClassificationFilter(new Set());
              setSelectedManager(null);
              setLineManagerFilter('all');
              setManagerTerritoryFilter('all');
              setSelectedRep(null);
              load(true);
            }}
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" style={{ display: 'inline-block', verticalAlign: 'middle' }}>
              <path d="M23 4v6h-6" /><path d="M1 20v-6h6" />
              <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15" />
            </svg>
          </button>
          <button className="hbtn hbtn-outline" onClick={signOut}>{t.signOut}</button>
        </div>
      </header>

      {/* LAYOUT: SIDEBAR + CONTENT */}
      <div className="dash-with-sidebar">

        {/* ── SIDEBAR ────────────────────────────────────────────── */}
        {tab !== 'roadmap' && (
          <aside className={`dash-sidebar${sidebarOpen ? ' open' : ''}`}>
            <button className="sb-close" onClick={() => setSidebarOpen(false)}>✕</button>

            {/* ─── SUMMARY SIDEBAR ─────────────────────────────── */}
            {tab === 'summary' && (
              <div className="sb-panel">
                {pmCoveragePieData.length > 0 && (
                  <>
                    <div className="sb-section-hd">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10" /><path d="M12 8v8M8 12h8" /></svg>
                      {rtl ? 'توزيع تغطية PM (عيادات vs مراكز)' : 'PM Coverage (Clinic vs Poly Clinic)'}
                    </div>
                    <PieChart data={pmCoveragePieData} title={rtl ? 'نسبة المساهمة %' : 'Contribution %'} />
                    <div className="sb-divider" />
                  </>
                )}
                {(activityEventData.totalActivities > 0 || activityEventData.totalEvents > 0) && (
                  <>
                    <div className="sb-section-hd">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 20V10M18 20V4M6 20v-4" /></svg>
                      {rtl ? 'الأنشطة والفعاليات' : 'Activities & Events'}
                    </div>
                    <div className="sb-team-group" style={{ marginTop: '8px', marginBottom: '16px' }}>
                      <div className="sb-rep-row" style={{ cursor: 'default', background: 'transparent' }}>
                        <span className="sb-rep-name">{t.kpi.no_activities || 'Activities'}</span>
                        <div className="sb-rep-bar-wrap">
                          <div className="sb-rep-bar" style={{ backgroundColor: '#8b5cf6', width: `${Math.min(100, activityEventData.totalActivities / Math.max(1, activityEventData.totalActivities, activityEventData.totalEvents) * 100)}%` }} />
                        </div>
                        <span className="sb-rep-val">{activityEventData.totalActivities}</span>
                      </div>
                      <div className="sb-rep-row" style={{ cursor: 'default', background: 'transparent' }}>
                        <span className="sb-rep-name">{t.kpi.no_events || 'Events'}</span>
                        <div className="sb-rep-bar-wrap">
                          <div className="sb-rep-bar" style={{ backgroundColor: '#ec4899', width: `${Math.min(100, activityEventData.totalEvents / Math.max(1, activityEventData.totalActivities, activityEventData.totalEvents) * 100)}%` }} />
                        </div>
                        <span className="sb-rep-val">{activityEventData.totalEvents}</span>
                      </div>
                    </div>
                    <div className="sb-divider" />
                  </>
                )}
                {selectedRepData ? (
                  <div className="sb-rep-detail">
                    <button className="sb-back" onClick={() => setSelectedRep(null)}>
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M19 12H5M12 19l-7-7 7-7" /></svg>
                      {rtl ? 'رجوع' : 'Back to list'}
                    </button>
                    <div className="sb-rep-hdr">
                      <div className="sb-rep-name-lg">{selectedRepData.user_name}</div>
                      <div className="sb-rep-team">{selectedRepData.team}{selectedRepData.is_manager ? ' · Manager' : ''}</div>
                      {selectedRepData.territory && <div className="sb-rep-terr">{selectedRepData.territory}</div>}
                    </div>
                    {t.kpiGroups.map(g => {
                      const keys = g.keys.filter(k => {
                        if (shift === 'AM') return !['pm_calls', 'pm_call_rate', 'pm_shift_days', 'total_pm_covered', 'clinic_covered', 'polyclinic_covered', 'avg_pm_shift_hm'].includes(k);
                        if (shift === 'PM') return !['am_calls', 'am_call_rate', 'am_shift_days', 'total_am_covered', 'amcenter_covered', 'hospital_covered', 'avg_am_shift_hm', 'avg_am_start_time'].includes(k);
                        return true;
                      });
                      // Hide entire group if no valid keys OR if it's the coaching tab and the user has no coaching records
                      if (g.keys.includes('coaching_days') && (selectedRepData.coaching_days === 0 || !selectedRepData.coaching_days)) return null;
                      const kpiRows = keys.map(k => ({ k, v: selectedRepData[k] })).filter(x => x.v !== null && x.v !== undefined && x.v !== '');
                      if (!kpiRows.length) return null;
                      return (
                        <div key={g.label} className="sb-kpi-sec">
                          <div className="sb-kpi-hd">{g.label}</div>
                          {kpiRows.map(({ k, v }) => (
                            <div key={k} className="sb-kpi-row">
                              <span>{k === 'coaching_days' ? (selectedRepData.is_manager ? t.kpi[k] : (rtl ? 'تم التوجيه' : 'Coached')) : (t.kpi[k] || k)}</span>
                              <span className="sb-kpi-val">{fmtVal(v, k)}</span>
                            </div>
                          ))}
                        </div>
                      );
                    })}
                  </div>
                ) : (
                  <>
                    <div className="sb-section-hd">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M23 21v-2a4 4 0 0 0-3-3.87" /><path d="M16 3.13a4 4 0 0 1 0 7.75" /></svg>
                      {rtl ? 'زيارات PM حسب المندوب' : 'PM Visits by Rep'}
                    </div>
                    {teamGroups.map(({ label, rows }) => {
                      const reps = [...rows].filter(r => !r.is_manager).sort((a, b) => (b.pm_calls || 0) - (a.pm_calls || 0));
                      if (!reps.length) return null;
                      return (
                        <div key={label} className="sb-team-group">
                          <div className="sb-team-label">{label}</div>
                          {reps.map(r => (
                            <div key={r.user_name}
                              className={`sb-rep-row${selectedRep === r.user_name ? ' active' : ''}`}
                              onClick={() => setSelectedRep(r.user_name)}>
                              <span className="sb-rep-name">{r.user_name}</span>
                              <div className="sb-rep-bar-wrap">
                                <div className="sb-rep-bar" style={{ width: `${Math.min(100, (r.pm_calls || 0) / Math.max(1, ...reps.map(x => x.pm_calls || 1)) * 100)}%` }} />
                              </div>
                              <span className="sb-rep-val">{r.pm_calls || 0}</span>
                            </div>
                          ))}
                        </div>
                      );
                    })}
                  </>
                )}
              </div>
            )}

            {/* ─── SPECIALTY SIDEBAR ───────────────────────────── */}
            {tab === 'specialty' && (
              <div className="sb-panel">
                {/* Specialty Dropdown Slicer */}
                <div className="sb-slicer" style={{ marginBottom: '24px' }}>
                  <div className="sb-slicer-hd">
                    <span>{rtl ? 'فلتر التخصص' : 'Filter Specialty'}</span>
                    {(specialtyFilter.size > 0 || classificationFilter.size > 0) && (
                      <button className="sb-slicer-clear" onClick={() => { setSpecialtyFilter(new Set()); setClassificationFilter(new Set()); }}>
                        {rtl ? 'مسح' : 'Clear'}
                      </button>
                    )}
                  </div>
                  <select
                    className="ctrl-sel"
                    style={{ width: '100%' }}
                    value={specialtyFilter.size === 1 ? Array.from(specialtyFilter)[0] : (specialtyFilter.size === 0 ? '' : 'mixed')}
                    onChange={(e) => {
                      const val = e.target.value;
                      if (val === '') setSpecialtyFilter(new Set());
                      else if (val !== 'mixed') setSpecialtyFilter(new Set([val]));
                    }}
                  >
                    <option value="">{rtl ? 'كل التخصصات' : 'All Specialties'}</option>
                    {specialtyFilter.size > 1 && <option value="mixed" disabled>{rtl ? 'تخصصات متعددة' : 'Multiple selected'}</option>}
                    {allSpecialties.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </div>

                <div className="sb-section-hd">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10" /><path d="M12 8v8M8 12h8" /></svg>
                  {rtl ? 'تغطية التخصصات' : 'Specialty Coverage'}
                </div>
                <PieChart
                  data={specialtyPieData}
                  title={rtl ? 'حسب التخصص' : 'By Specialty'}
                  onSelect={(s) => {
                    setSpecialtyFilter(prev => {
                      if (prev.has(s) && prev.size === 1) return new Set();
                      return new Set([s]);
                    });
                  }}
                  activeFilters={specialtyFilter}
                />

                <div className="sb-divider" />

                <PieChart
                  data={classificationPieData}
                  title={rtl ? 'حسب التصنيف' : 'By Classification'}
                  onSelect={(c) => {
                    setClassificationFilter(prev => {
                      if (prev.has(c) && prev.size === 1) return new Set();
                      return new Set([c]);
                    });
                  }}
                  activeFilters={classificationFilter}
                />
              </div>
            )}

            {/* ─── PRODUCTS SIDEBAR ────────────────────────────── */}
            {tab === 'products' && (
              <div className="sb-panel">
                {/* Product Dropdown Slicer — matches the Specialty slicer pattern */}
                <div className="sb-slicer" style={{ marginBottom: '24px' }}>
                  <div className="sb-slicer-hd">
                    <span>{rtl ? 'فلتر المنتج' : 'Filter Product'}</span>
                    {productFilter.size > 0 && (
                      <button className="sb-slicer-clear" onClick={() => setProductFilter(new Set())}>
                        {rtl ? 'مسح' : 'Clear'}
                      </button>
                    )}
                  </div>
                  <select
                    className="ctrl-sel"
                    style={{ width: '100%' }}
                    value={productFilter.size === 1 ? Array.from(productFilter)[0] : (productFilter.size === 0 ? '' : 'mixed')}
                    onChange={(e) => {
                      const val = e.target.value;
                      if (val === '') setProductFilter(new Set());
                      else if (val !== 'mixed') setProductFilter(new Set([val]));
                    }}
                  >
                    <option value="">{rtl ? 'كل المنتجات' : 'All Products'}</option>
                    {productFilter.size > 1 && <option value="mixed" disabled>{rtl ? 'منتجات متعددة' : 'Multiple selected'}</option>}
                    {allProducts.map(p => (
                      <option key={p} value={p}>{p}</option>
                    ))}
                  </select>
                </div>

                <div className="sb-section-hd">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="7" height="7" /><rect x="14" y="3" width="7" height="7" /><rect x="14" y="14" width="7" height="7" /><rect x="3" y="14" width="7" height="7" /></svg>
                  {rtl ? 'مساهمة المنتجات' : 'Product Contribution'}
                </div>
                <PieChart
                  data={productPieData}
                  title={rtl ? 'حسب المنتج' : 'By Product'}
                  onSelect={(p) => {
                    setProductFilter(prev => {
                      if (prev.has(p) && prev.size === 1) return new Set();
                      return new Set([p]);
                    });
                  }}
                  activeFilters={productFilter}
                />

                <div className="sb-divider" />

                <div className="sb-section-hd" style={{ marginTop: 0 }}>
                  {rtl ? 'أعلى المنتجات' : 'Top Products'}
                </div>
                <div className="sb-top-list">
                  {topProducts.map((p, i) => (
                    <div key={p.name} className="sb-top-item">
                      <span className="sb-top-rank">#{i + 1}</span>
                      <div className="sb-top-info">
                        <div className="sb-top-name">{p.name}</div>
                        <div className="sb-top-bar-wrap">
                          <div className="sb-top-bar" style={{ width: `${p.pct}%` }} />
                        </div>
                      </div>
                      <span className="sb-top-count">{p.count}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* ─── COACHING SIDEBAR ────────────────────────────── */}
            {tab === 'coaching' && (
              <div className="sb-panel">
                <div className="sb-section-hd">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /></svg>
                  {rtl ? 'المديرون' : 'Managers'}
                </div>
                {selectedManager && (
                  <button className="sb-back" onClick={() => setSelectedManager(null)}>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M19 12H5M12 19l-7-7 7-7" /></svg>
                    {rtl ? 'عرض الكل' : 'Show all'}
                  </button>
                )}
                {managerGroups.length > 0 ? managerGroups.map(mgr => (
                  <div key={mgr.name}
                    className={`sb-mgr-card${selectedManager === mgr.name ? ' active' : ''}`}
                    onClick={() => setSelectedManager(selectedManager === mgr.name ? null : mgr.name)}>
                    <div className="sb-mgr-avatar">
                      {mgr.name.split(' ').map(w => w[0]).slice(0, 2).join('')}
                    </div>
                    <div className="sb-mgr-info">
                      <div className="sb-mgr-name">{mgr.name}</div>
                      <div className="sb-mgr-meta">{mgr.team}</div>
                    </div>
                    <div className="sb-mgr-stats">
                      <div className="sb-mgr-stat">{mgr.dayCount}<small> days</small></div>
                      <div className="sb-mgr-stat">{mgr.repCount}<small> reps</small></div>
                    </div>
                  </div>
                )) : (
                  <div style={{ padding: '16px', color: 'var(--text-light)', fontSize: '13px', textAlign: 'center' }}>
                    {rtl ? 'لا يوجد مديرون لهذه الفترة.' : 'No managers for this period.'}
                  </div>
                )}
              </div>
            )}

            {/* ─── TIMING SIDEBAR ──────────────────────────────── */}
            {tab === 'timing' && (
              <div className="sb-panel">
                <div className="sb-section-hd">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10" /><path d="M12 6v6l4 2" /></svg>
                  {rtl ? 'توزيع الزيارات الأخيرة' : 'Last Visit Distribution'}
                </div>
                <PieChart
                  data={[
                    { label: t.kpi.timing_early || 'Before 3 PM', value: timingStats.early, color: '#ef4444' },
                    { label: t.kpi.timing_normal || '3 PM – 6 PM', value: timingStats.normal, color: '#10b981' },
                    { label: t.kpi.timing_late || 'After 6 PM', value: timingStats.late, color: '#f59e0b' },
                  ].filter(d => d.value > 0)}
                  title={rtl ? 'التوزيع' : 'Distribution'}
                  onSelect={(label) => {
                    const cat = label === (t.kpi.timing_early || 'Before 3 PM') ? 'early'
                      : label === (t.kpi.timing_normal || '3 PM – 6 PM') ? 'normal'
                        : label === (t.kpi.timing_late || 'After 6 PM') ? 'late' : 'all';
                    setTimingCategoryFilter(prev => prev === cat ? 'all' : cat);
                  }}
                />
                <div className="sb-divider" />
                {/* Per-rep summary */}
                <div className="sb-section-hd" style={{ marginTop: 0 }}>
                  {rtl ? 'ملخص حسب المندوب' : 'Per-Rep Summary'}
                </div>
                <div className="sb-top-list">
                  {(() => {
                    const byRep = {};
                    fTiming.forEach(r => {
                      if (!byRep[r.user]) byRep[r.user] = { early: 0, normal: 0, late: 0, total: 0 };
                      byRep[r.user][r.category] = (byRep[r.user][r.category] || 0) + 1;
                      byRep[r.user].total++;
                    });
                    return Object.entries(byRep).sort((a, b) => b[1].early - a[1].early).slice(0, 15).map(([rep, stats]) => (
                      <div key={rep} className="sb-top-item" style={{ cursor: 'pointer' }}
                        onClick={() => { setUser(rep === userFilter ? 'all' : rep); setSelectedRep(rep === userFilter ? null : rep); }}>
                        <div className="sb-top-info" style={{ flex: 1 }}>
                          <div className="sb-top-name">{rep}</div>
                          <div className="timing-rep-bars">
                            {stats.early > 0 && <span className="timing-mini-badge timing-badge-early">{stats.early}</span>}
                            {stats.normal > 0 && <span className="timing-mini-badge timing-badge-normal">{stats.normal}</span>}
                            {stats.late > 0 && <span className="timing-mini-badge timing-badge-late">{stats.late}</span>}
                          </div>
                        </div>
                        <span className="sb-top-count">{stats.total}d</span>
                      </div>
                    ));
                  })()}
                </div>
              </div>
            )}
          </aside>
        )}

        {/* Mobile backdrop */}
        {sidebarOpen && <div className="sb-backdrop" onClick={() => setSidebarOpen(false)} />}

        {/* ── MAIN CONTENT ───────────────────────────────────── */}
        <div className="dash-content">

          {/* CONTROL BAR */}
          <div className="ctrl-bar">
            <div className="ctrl-row">
              <div className="ctrl-group">
                <span className="ctrl-lbl">{rtl ? 'الفترة' : 'Period'}</span>
                <select className="ctrl-sel" value={period} onChange={e => setPeriod(e.target.value)}>
                  {availablePeriods.map(p => (
                    <option key={p} value={p}>{p}</option>
                  ))}
                </select>
              </div>
              <div className="ctrl-group">
                <span className="ctrl-lbl">{rtl ? 'النطاق الزمني' : 'Time Slicer'}</span>
                <select className="ctrl-sel" value={timeGrain} onChange={e => setTimeGrain(e.target.value)}>
                  <option value="all">{rtl ? 'الشهر كامل' : 'Full Month'}</option>
                  <option value="biweekly1">{rtl ? 'النصف الأول (1–15)' : 'Bi-Weekly (1–15)'}</option>
                  <option value="biweekly2">{rtl ? 'النصف الثاني (16–31)' : 'Bi-Weekly (16–31)'}</option>
                  <option value="week1">{rtl ? 'الأسبوع الأول (1–7)' : 'Week 1 (1–7)'}</option>
                  <option value="week2">{rtl ? 'الأسبوع الثاني (8–14)' : 'Week 2 (8–14)'}</option>
                  <option value="week3">{rtl ? 'الأسبوع الثالث (15–21)' : 'Week 3 (15–21)'}</option>
                  <option value="week4">{rtl ? 'الأسبوع الرابع (22–31)' : 'Week 4 (22–31)'}</option>
                </select>
              </div>
              <div className="ctrl-group">
                <span className="ctrl-lbl">{rtl ? 'الوردية' : 'Shift'}</span>
                <ShiftToggle value={shift} onChange={setShift} t={t} />
              </div>
              {teams.length > 1 && (
                <div className="ctrl-group">
                  <span className="ctrl-lbl">{rtl ? 'الفريق' : 'Team'}</span>
                  <select className="ctrl-sel" value={team} onChange={e => { setTeam(e.target.value); setUser('all'); }}>
                    <option value="all">{t.allTeams}</option>
                    {teams.map(tm => <option key={tm} value={tm}>{tm}</option>)}
                  </select>
                </div>
              )}
              {allLineManagers.length > 0 && (
                <div className="ctrl-group">
                  <span className="ctrl-lbl">{rtl ? 'مدير الخط' : 'Line Manager'}</span>
                  <select className="ctrl-sel" value={lineManagerFilter} onChange={e => { setLineManagerFilter(e.target.value); setUser('all'); }}>
                    <option value="all">{rtl ? 'كل مديري الخطوط' : 'All Line Managers'}</option>
                    {allLineManagers.map(lm => <option key={lm} value={lm}>{lm}</option>)}
                  </select>
                </div>
              )}
              {allManagerTerritories.length > 0 && (
                <div className="ctrl-group">
                  <span className="ctrl-lbl">{rtl ? 'منطقة المدير' : 'Manager Territory'}</span>
                  <select className="ctrl-sel" value={managerTerritoryFilter} onChange={e => setManagerTerritoryFilter(e.target.value)}>
                    <option value="all">{rtl ? 'كل المناطق' : 'All Territories'}</option>
                    {allManagerTerritories.map(mt => <option key={mt} value={mt}>{mt}</option>)}
                  </select>
                </div>
              )}
              {isMgr && allUsers.length > 1 && (
                <div className="ctrl-group">
                  <span className="ctrl-lbl">{rtl ? 'المندوب' : 'Rep'}</span>
                  <select className="ctrl-sel" value={userFilter} onChange={e => {
                    const val = e.target.value;
                    setUser(val);
                    setSelectedRep(val === 'all' ? null : val);
                  }}>
                    <option value="all">{t.allUsers}</option>
                    {allUsers.map(u => <option key={u} value={u}>{u}</option>)}
                  </select>
                </div>
              )}
              {tab === 'summary' && (
                <div className="ctrl-group">
                  <div className="shift-toggle">
                    <button className={`stoggle${viewMode === 'teams' ? ' on' : ''}`} onClick={() => setViewMode('teams')}>
                      👥 {rtl ? 'ملخص الفرق' : 'Team Brief'}
                    </button>
                    <button className={`stoggle${viewMode === 'employees' ? ' on' : ''}`} onClick={() => setViewMode('employees')}>
                      👤 {rtl ? 'المندوبون' : 'Employee Brief'}
                    </button>
                  </div>
                </div>
              )}
              <div className="ctrl-group ctrl-search">
                <div className="search-box">
                  <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="2">
                    <circle cx="9" cy="9" r="6" /><path d="M15 15l-3.5-3.5" />
                  </svg>
                  <input className="search-inp" placeholder={t.search}
                    value={search} onChange={e => setSearch(e.target.value)} />
                  {search && <button className="search-clear" onClick={() => setSearch('')}>✕</button>}
                </div>
              </div>
              <div className="ctrl-end">
                <button className="hbtn hbtn-primary" onClick={doExport}>↓ Export — Be The Analyst</button>
              </div>
            </div>
          </div>

          {/* TABS */}
          <nav className="dash-tabs">
            {tab !== 'roadmap' && (
              <button className="sidebar-toggle" onClick={() => setSidebarOpen(!sidebarOpen)} title="Toggle panel">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <rect x="3" y="3" width="18" height="18" rx="2" /><path d="M9 3v18" />
                </svg>
              </button>
            )}
            {visibleTabs.map(([k, label]) => (
              <button key={k} className={`dtab${tab === k ? ' on' : ''}`} onClick={() => changeTab(k)}>{label}</button>
            ))}
          </nav>

          {error && <div className="dash-err">{error}</div>}

          {loading ? (
            <div className="dash-empty">{t.loading}</div>
          ) : (
            <div className="dash-body">

              {/* SUMMARY TAB */}
              {tab === 'summary' && (
                fSummary.length === 0 ? <div className="dash-empty">{t.noData}</div> : (
                  <>
                    {viewMode === 'teams' ? (
                      <div className="cards-grid">
                        {teamGroups.map(({ label, rows }) => (
                          <TeamBriefCard
                            key={label}
                            rows={rows}
                            teamLabel={label}
                            rtl={rtl}
                            t={t}
                            shift={shift}
                            isMgr={isMgr}
                            onSelectTeam={(tName) => {
                              setTeam(tName);
                              setViewMode('employees');
                            }}
                          />
                        ))}
                      </div>
                    ) : (
                      <div className="cards-grid">
                        {fSummary.map((r, i) => {
                          const normUser = (r.user_name || '').toLowerCase().trim();
                          const rawRole = r.role || userHierarchyMap[normUser]?.role || (r.is_manager ? 'Supervisor' : 'MR');
                          const roleLower = String(rawRole).toLowerCase();
                          const roleClass = roleLower.includes('area') ? 'hdr-role-am'
                            : (roleLower.includes('supervisor') || roleLower.includes('sup')) ? 'hdr-role-sup'
                              : roleLower.includes('blm') ? 'hdr-role-blm'
                                : 'hdr-role-mr';

                          const roleLabel = roleLower.includes('area') ? (rtl ? 'مدير منطقة' : 'Area Manager')
                            : (roleLower.includes('supervisor') || roleLower.includes('sup')) ? (rtl ? 'مشرف' : 'Supervisor')
                              : roleLower.includes('blm') ? (rtl ? 'مدير خط' : 'BLM')
                                : (rtl ? 'مندوب' : 'MR');

                          return (
                            <div key={r.id || i} className={`ucard ${roleClass}${r.is_manager ? ' mgr' : ''}${selectedRep === r.user_name ? ' ucard-selected' : ''}`}
                              onClick={() => handleSelectRep(r.user_name)}>
                              <div className={`ucard-hdr ${roleClass}`}>
                                <div className="ucard-info">
                                  <div className="ucard-name">{r.user_name}</div>
                                  <div className="ucard-meta">{roleLower.includes('mr') ? (r.team ? `${r.team} · ${roleLabel}` : roleLabel) : (r.team || '')}</div>
                                  {r.territory && <div className="ucard-terr" title={r.territory}>{r.territory}</div>}
                                  {(r.avg_am_shift_hm || r.avg_pm_shift_hm) && (
                                    <div className="ucard-dur">
                                      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><circle cx="12" cy="12" r="10" /><path d="M12 6v6l4 2" /></svg>
                                      {r.avg_am_shift_hm ? <span className="dur-am">AM {fmtDuration(r.avg_am_shift_hm)}</span> : null}
                                      {r.avg_pm_shift_hm ? <span className="dur-pm">PM {fmtDuration(r.avg_pm_shift_hm)}</span> : null}
                                    </div>
                                  )}
                                </div>
                                <span className={`mgr-pip ${roleClass}`}>{roleLabel.toUpperCase()}</span>
                              </div>
                              {t.kpiGroups.map(g => {
                                const keys = g.keys.filter(k => {
                                  if (shift === 'AM') return !['pm_calls', 'pm_call_rate', 'pm_shift_days', 'total_pm_covered', 'clinic_covered', 'polyclinic_covered', 'avg_pm_shift_hm'].includes(k);
                                  if (shift === 'PM') return !['am_calls', 'am_call_rate', 'am_shift_days', 'total_am_covered', 'amcenter_covered', 'hospital_covered', 'avg_am_shift_hm', 'avg_am_start_time'].includes(k);
                                  return true;
                                });
                                if (g.keys.includes('coaching_days') && !isMgr) return null;
                                const kpiRows = keys.map(k => ({ k, v: r[k] })).filter(x => x.v !== null && x.v !== undefined && x.v !== '');
                                if (!kpiRows.length) return null;
                                return (
                                  <div key={g.label} className={`kpi-sec${g.keys.includes('avg_am_start_time') ? ' kpi-timing' : ''}`}>
                                    <div className="kpi-sec-hd">{g.label}</div>
                                    {kpiRows.map(({ k, v }) => {
                                      const target = KPI_TARGETS[k];
                                      const numVal = Number(v) || 0;
                                      const pct = target ? Math.min(100, Math.round((numVal / target) * 100)) : null;
                                      return (
                                        <div key={k} className="kpi-row-wrapper">
                                          <div className="kpi-row">
                                            <span className="kpi-lbl">{t.kpi[k] || k}</span>
                                            <span className={`kpi-v ${k.includes('rate') ? 'rate' : ''}`}>{fmtVal(v, k)}</span>
                                          </div>
                                          {pct !== null && (
                                            <div className="kpi-card-progress" title={`${pct}% of target (${target})`}>
                                              <div className="kpi-card-progress-bar" style={{ width: `${pct}%`, backgroundColor: pct >= 100 ? '#10b981' : pct >= 70 ? '#3b82f6' : '#ef4444' }} />
                                            </div>
                                          )}
                                        </div>
                                      );
                                    })}
                                  </div>
                                );
                              })}
                              {r.product_calls_detail && shift !== 'AM' && (
                                <div className="kpi-sec">
                                  <div className="kpi-sec-hd">{rtl ? 'تفاصيل المنتج' : 'Product Detail'}</div>
                                  <div className="prod-det">{r.product_calls_detail}</div>
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </>
                )
              )}

              {/* SPECIALTY TAB */}
              {tab === 'specialty' && (
                <>
                  <PivotSummaryBanner
                    rows={fSpecialty}
                    valueKey="call_count"
                    rowKey="specialty"
                    shift={shift}
                    t={t}
                    selectedTeam={team}
                    onSelectTeam={setTeam}
                    userTeamMap={userTeamMap}
                  />
                  <PivotTable rows={fSpecialty} rowKey="specialty" valueKey="call_count"
                    shiftFilter={shift} userFilter={userFilter} searchFilter={search} lang={lang} hideAvg={true} />
                </>
              )}

              {/* PRODUCTS TAB */}
              {tab === 'products' && (
                <>
                  <PivotSummaryBanner
                    rows={fProducts}
                    valueKey="call_count"
                    rowKey="product"
                    shift={shift}
                    t={t}
                    selectedTeam={team}
                    onSelectTeam={setTeam}
                    userTeamMap={userTeamMap}
                  />
                  <PivotTable rows={fProducts} rowKey="product" valueKey="call_count"
                    shiftFilter={shift} userFilter={userFilter} searchFilter={search} lang={lang} />
                </>
              )}

              {/* COACHING TAB */}
              {tab === 'coaching' && isMgr && (
                filteredCoaching.length === 0 ? (
                  <div className="dash-empty">
                    {selectedManager
                      ? (rtl ? 'لا توجد بيانات لهذا المدير' : 'No coaching data for this manager')
                      : t.noData}
                  </div>
                ) : (
                  <>
                    {selectedManager && (
                      <div className="coaching-selected-hdr">
                        <span className="coaching-sel-name">{selectedManager}</span>
                        <span className="coaching-sel-meta">
                          {filteredCoaching.length} {rtl ? 'جلسة' : 'session'}{filteredCoaching.length !== 1 ? 's' : ''}
                        </span>
                      </div>
                    )}
                    <div className="pivot-wrap">
                      <table className="pivot-tbl">
                        <thead>
                          <tr>
                            <th className="s-col">{rtl ? 'المدير' : 'Manager'}</th>
                            <th className="rep-col">{rtl ? 'المندوب' : 'Rep'}</th>
                            <th>{rtl ? 'التاريخ' : 'Date'}</th>
                            <th>{rtl ? 'الفريق' : 'Team'}</th>
                            <th>{rtl ? 'زيارات AM' : 'AM Visits'}</th>
                            <th>{rtl ? 'مرافقة AM' : 'AM Acc.'}</th>
                            <th>AM %</th>
                            <th>{rtl ? 'زيارات PM' : 'PM Visits'}</th>
                            <th>{rtl ? 'مرافقة PM' : 'PM Acc.'}</th>
                            <th>PM %</th>
                          </tr>
                        </thead>
                        <tbody>
                          {[...filteredCoaching].sort((a, b) => (a.manager_name || '').localeCompare(b.manager_name || '') || (a.coaching_date || '').localeCompare(b.coaching_date || '')).map((r, i) => (
                            <tr key={r.id || i}>
                              <td className="s-col">{r.manager_name}</td>
                              <td className="rep-col">{r.rep_name}</td>
                              <td>{r.coaching_date}</td>
                              <td>{r.team || '—'}</td>
                              <td>{r.am_visits || 0}</td>
                              <td>{r.am_accompanied || 0}</td>
                              <td>{r.am_visits ? Math.round((r.am_accompanied / r.am_visits) * 100) + '%' : '-'}</td>
                              <td>{r.pm_visits || 0}</td>
                              <td>{r.pm_accompanied || 0}</td>
                              <td>{r.pm_visits ? Math.round((r.pm_accompanied / r.pm_visits) * 100) + '%' : '-'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </>
                )
              )}

              {/* TIMING TAB */}
              {tab === 'timing' && (
                filteredTiming.length === 0 ? (
                  <div className="dash-empty">{t.noData}</div>
                ) : (
                  <>
                    {/* Timing Summary Stats */}
                    <div className="timing-stats-banner">
                      <div className={`timing-stat-card${timingCategoryFilter === 'all' ? ' active' : ''}`}
                        onClick={() => setTimingCategoryFilter('all')}>
                        <div className="timing-stat-val">{timingStats.uniqueDays}</div>
                        <div className="timing-stat-lbl">{rtl ? 'أيام العمل' : 'Working Days'}</div>
                        <div className="timing-stat-pct">{timingStats.total} {rtl ? 'سجل' : 'records'}</div>
                      </div>
                      <div className={`timing-stat-card timing-early${timingCategoryFilter === 'early' ? ' active' : ''}`}
                        onClick={() => setTimingCategoryFilter(timingCategoryFilter === 'early' ? 'all' : 'early')}>
                        <div className="timing-stat-val">{timingStats.early}</div>
                        <div className="timing-stat-lbl">{t.kpi.timing_early || 'Before 3 PM'}</div>
                        <div className="timing-stat-pct">{timingStats.total ? Math.round(timingStats.early / timingStats.total * 100) : 0}%</div>
                      </div>
                      <div className={`timing-stat-card timing-normal${timingCategoryFilter === 'normal' ? ' active' : ''}`}
                        onClick={() => setTimingCategoryFilter(timingCategoryFilter === 'normal' ? 'all' : 'normal')}>
                        <div className="timing-stat-val">{timingStats.normal}</div>
                        <div className="timing-stat-lbl">{t.kpi.timing_normal || '3 PM – 6 PM'}</div>
                        <div className="timing-stat-pct">{timingStats.total ? Math.round(timingStats.normal / timingStats.total * 100) : 0}%</div>
                      </div>
                      <div className={`timing-stat-card timing-late${timingCategoryFilter === 'late' ? ' active' : ''}`}
                        onClick={() => setTimingCategoryFilter(timingCategoryFilter === 'late' ? 'all' : 'late')}>
                        <div className="timing-stat-val">{timingStats.late}</div>
                        <div className="timing-stat-lbl">{t.kpi.timing_late || 'After 6 PM'}</div>
                        <div className="timing-stat-pct">{timingStats.total ? Math.round(timingStats.late / timingStats.total * 100) : 0}%</div>
                      </div>
                    </div>

                    {/* Timing Detail Table */}
                    {userFilter === 'all' && !selectedRep ? (
                      <div className="timing-prompt-card">
                        <div className="timing-prompt-icon">👤</div>
                        <div className="timing-prompt-title">
                          {rtl ? 'اختر مندوباً لعرض سجل الزيارات الأخيرة' : 'Select an Employee to View Last Visit Log'}
                        </div>
                        <div className="timing-prompt-desc">
                          {rtl ? 'تم عرض إحصائيات ونسب الفئات أعلاه. يرجى اختيار مندوب محدد من القائمة أو الشريط الجانبي لعرض سجل زياراته اليومي.'
                            : 'High-level category insights are summarized above. Select a specific employee from the REP dropdown or sidebar to inspect their exact daily last visit records.'}
                        </div>
                      </div>
                    ) : (
                      <>
                        <div className="pivot-wrap">
                          <table className="pivot-tbl timing-tbl">
                            <thead>
                              <tr>
                                <th className="s-col">{rtl ? 'التاريخ' : 'Date'}</th>
                                <th>{rtl ? 'المندوب' : 'Rep'}</th>
                                <th>{rtl ? 'الفريق' : 'Team'}</th>
                                <th>{rtl ? 'آخر زيارة' : 'Last Visit'}</th>
                                <th>{rtl ? 'الفئة' : 'Category'}</th>
                              </tr>
                            </thead>
                            <tbody>
                              {filteredTiming.slice(0, timingDisplayLimit).map((r, i) => (
                                <tr key={`${r.user}-${r.date}-${i}`} className={`timing-row timing-row-${r.category}`}>
                                  <td className="s-col">{r.date}</td>
                                  <td>{r.user}</td>
                                  <td>{r.team || '—'}</td>
                                  <td className="timing-time">{r.formattedTime}</td>
                                  <td>
                                    <span className={`timing-badge timing-badge-${r.category}`}>
                                      {r.category === 'early' ? (t.kpi.timing_early || '< 3 PM')
                                        : r.category === 'normal' ? (t.kpi.timing_normal || '3–6 PM')
                                          : r.category === 'late' ? (t.kpi.timing_late || '> 6 PM')
                                            : '—'}
                                    </span>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                        {filteredTiming.length > timingDisplayLimit && (
                          <div className="timing-load-more">
                            <button className="timing-btn-more" onClick={() => setTimingDisplayLimit(prev => prev + 200)}>
                              {rtl ? `عرض المزيد (عرض ${timingDisplayLimit} من إجمالي ${filteredTiming.length})` : `Show More (Showing ${timingDisplayLimit} of ${filteredTiming.length} records)`}
                            </button>
                            <button className="timing-btn-all" onClick={() => setTimingDisplayLimit(filteredTiming.length)}>
                              {rtl ? 'عرض الكل' : 'Show All'}
                            </button>
                          </div>
                        )}
                      </>
                    )}
                  </>
                )
              )}
            </div>
          )}
        </div>
      </div>

      {/* AI Chat UI */}
      <div className={`ai-chat-widget ${isAiOpen ? 'open' : ''}`}>
        {!isAiOpen && (
          <button className="ai-fab" onClick={() => setIsAiOpen(true)}>
            <span className="ai-fab-icon">✨</span> Ask AI
          </button>
        )}
        {isAiOpen && (
          <div className="ai-chat-window">
            <div className="ai-chat-header">
              <span className="ai-chat-title">✨ Excellence AI</span>
              <button className="ai-chat-close" onClick={() => setIsAiOpen(false)}>✕</button>
            </div>
            <div className="ai-chat-body">
              {aiHistory.length === 0 ? (
                <div className="ai-welcome">
                  Ask me anything about the current dashboard data!
                  <br /><br />
                  <small>Example: "Who are the top performers in EAGLES 1?"</small>
                </div>
              ) : (
                aiHistory.map((msg, idx) => (
                  <div key={idx} className={`ai-msg ${msg.role}`}>
                    {msg.content}
                  </div>
                ))
              )}
              {isAiLoading && <div className="ai-msg assistant loading">Thinking...</div>}
            </div>
            <form className="ai-chat-input-area" onSubmit={handleAiSubmit}>
              <input
                type="text"
                placeholder="Ask about this data..."
                value={aiInput}
                onChange={e => setAiInput(e.target.value)}
                disabled={isAiLoading}
              />
              <button type="submit" disabled={isAiLoading || !aiInput.trim()}>Send</button>
            </form>
          </div>
        )}
      </div>

    </div>
  );
}