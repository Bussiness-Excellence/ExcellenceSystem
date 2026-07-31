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
    sliceUnavailable: 'Visit-level data is not loaded for this period, so date-range slicing is unavailable. Showing Full Month instead.',
    sliceNoDates: 'This view has no per-visit dates available, so it cannot be filtered by date range.',
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
    sliceUnavailable: 'بيانات الزيارات غير محمّلة لهذه الفترة، لذا لا يمكن التصفية حسب النطاق الزمني. يتم عرض الشهر بالكامل.',
    sliceNoDates: 'لا تتوفر تواريخ زيارات لهذا العرض، لذا لا يمكن تصفيته حسب النطاق الزمني.',
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

function formatDateDisplay(val) {
  if (!val) return '—';
  const s = String(val).trim();
  // Period format "YYYY-MM" or "2026-07"
  const mPeriod = s.match(/^(\d{4})[-/](\d{1,2})$/);
  if (mPeriod) {
    const yr = parseInt(mPeriod[1], 10);
    const mo = parseInt(mPeriod[2], 10) - 1;
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    if (mo >= 0 && mo < 12) return `${monthNames[mo]} ${yr}`;
  }
  // Date format "YYYY-MM-DD"
  const mIso = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (mIso) {
    const yr = parseInt(mIso[1], 10);
    const mo = parseInt(mIso[2], 10) - 1;
    const da = parseInt(mIso[3], 10);
    const shortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    if (mo >= 0 && mo < 12) return `${shortMonths[mo]} ${da}, ${yr}`;
  }
  // Date format "DD-MM-YYYY" or "MM/DD/YYYY"
  const mUs = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (mUs) {
    const p1 = parseInt(mUs[1], 10);
    const p2 = parseInt(mUs[2], 10);
    const yr = parseInt(mUs[3], 10);
    const isMMDDYYYY = p1 <= 12 && p2 > 12;
    const mo = (isMMDDYYYY ? p1 : p2) - 1;
    const da = isMMDDYYYY ? p2 : p1;
    const shortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    if (mo >= 0 && mo < 12) return `${shortMonths[mo]} ${da}, ${yr}`;
  }
  return s;
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
                  const useAvg = k.includes('rate') || k.includes('avg_');
                  displayVal = fmtVal(useAvg ? item.avg : item.sum, k);
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

  const grandTotal = useMemo(() => filtered.reduce((s, r) => s + (r[valueKey] || 0), 0), [filtered, valueKey]);
  const activeRepsCount = useMemo(() => new Set(filtered.map(r => r.user_name).filter(Boolean)).size, [filtered]);

  const byTeam = useMemo(() => {
    const m = {};
    filtered.forEach(r => {
      const rawTeam = (r.team && r.team !== 'Unknown') ? r.team : (userTeamMap && userTeamMap[r.user_name]) || 'Other';
      const primaryTeam = (typeof rawTeam === 'string') ? rawTeam.split('; ')[0].trim() : rawTeam;
      if (!m[primaryTeam]) m[primaryTeam] = { total: 0, users: new Set() };
      m[primaryTeam].total += (r[valueKey] || 0);
      m[primaryTeam].users.add(r.user_name);
    });
    return m;
  }, [filtered, valueKey, userTeamMap]);

  const teamList = Object.entries(byTeam).sort((a, b) => b[1].total - a[1].total);
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
        <span className="pb-sub">{activeRepsCount} active reps</span>
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
function PivotTable({ rows, rowKey, colKey = 'user_name', valueKey, secondValueKey, shiftFilter, userFilter, searchFilter, lang, hideAvg, rowTitle, colTitle }) {
  const filtered = useMemo(() => rows.filter(r => {
    if (shiftFilter !== 'all' && r.shift !== shiftFilter) return false;
    if (userFilter && userFilter !== 'all' && r.user_name !== userFilter) return false;
    if (searchFilter && !r[rowKey]?.toLowerCase().includes(searchFilter.toLowerCase())
      && !r.user_name?.toLowerCase().includes(searchFilter.toLowerCase())) return false;
    return true;
  }), [rows, rowKey, shiftFilter, userFilter, searchFilter]);

  const cols = useMemo(() => [...new Set(filtered.map(r => r[colKey] || 'Unassigned'))].sort(), [filtered, colKey]);
  const rowKeys = useMemo(() => [...new Set(filtered.map(r => r[rowKey] || 'Other'))].sort(), [filtered, rowKey]);

  const cells = useMemo(() => {
    const c = {};
    filtered.forEach(r => {
      const rVal = r[rowKey] || 'Other';
      const cVal = r[colKey] || 'Unassigned';
      if (!c[rVal]) c[rVal] = {};
      if (!c[rVal][cVal]) c[rVal][cVal] = { v1: 0, v2: 0 };
      c[rVal][cVal].v1 += (r[valueKey] || 0);
      if (secondValueKey) c[rVal][cVal].v2 += (r[secondValueKey] || 0);
    });
    return c;
  }, [filtered, rowKey, colKey, valueKey, secondValueKey]);

  const colTotals = useMemo(() => {
    const ct = {};
    cols.forEach(col => {
      const colRows = filtered.filter(r => (r[colKey] || 'Unassigned') === col);
      ct[col] = {
        v1: colRows.reduce((s, r) => s + (r[valueKey] || 0), 0),
        v2: secondValueKey ? colRows.reduce((s, r) => s + (r[secondValueKey] || 0), 0) : 0
      };
    });
    return ct;
  }, [cols, filtered, colKey, valueKey, secondValueKey]);

  // Synced top scrollbar
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
  }, [filtered, cols, rowKeys]);

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

  const defaultRowTitle = rowKey === 'specialty' ? (lang === 'ar' ? 'التخصص' : 'Specialty') : rowKey === 'product' ? (lang === 'ar' ? 'المنتج' : 'Product') : rowKey;

  return (
    <>
      <div className="pivot-top-scroll" ref={topScrollRef} onScroll={handleTopScroll}>
        <div style={{ width: tableWidth, height: 1 }} />
      </div>
      <div className="pivot-wrap" ref={wrapRef} onScroll={handleWrapScroll}>
        <table className="pivot-tbl">
          <thead>
            <tr>
              <th className="s-col">{rowTitle || defaultRowTitle}</th>
              {cols.map(c => <th key={c} title={c}>{colKey === 'user_name' ? c.split(' ').slice(0, 2).join(' ') : c}</th>)}
              <th className="t-col">Σ Total</th>
            </tr>
            {!hideAvg && (
              <tr className="avg-row">
                <th className="s-col avg-lbl">⌀ Avg</th>
                {cols.map(c => {
                  const cTotal = colTotals[c]?.v1 || 0;
                  const cRows = rowKeys.filter(k => cells[k]?.[c]?.v1).length;
                  const avgV1 = cRows > 0 ? Math.round(cTotal / cRows) : 0;
                  if (secondValueKey) {
                    const cTotal2 = colTotals[c]?.v2 || 0;
                    const avgV2 = cRows > 0 ? Math.round(cTotal2 / cRows) : 0;
                    return <th key={c} className="avg-cell">{avgV1} / {avgV2}</th>;
                  }
                  return <th key={c} className="avg-cell">{avgV1}</th>;
                })}
                <th className="t-col avg-cell">
                  {(() => {
                    const gt1 = filtered.reduce((s, r) => s + (r[valueKey] || 0), 0);
                    const avgGt1 = cols.length > 0 ? Math.round(gt1 / cols.length) : 0;
                    if (secondValueKey) {
                      const gt2 = filtered.reduce((s, r) => s + (r[secondValueKey] || 0), 0);
                      const avgGt2 = cols.length > 0 ? Math.round(gt2 / cols.length) : 0;
                      return `${avgGt1} / ${avgGt2}`;
                    }
                    return avgGt1;
                  })()}
                </th>
              </tr>
            )}
          </thead>
          <tbody>
            {rowKeys.map(k => {
              const rowTotal1 = cols.reduce((s, c) => s + (cells[k]?.[c]?.v1 || 0), 0);
              const rowTotal2 = secondValueKey ? cols.reduce((s, c) => s + (cells[k]?.[c]?.v2 || 0), 0) : 0;
              return (
                <tr key={k}>
                  <td className="s-col">{k}</td>
                  {cols.map(c => {
                    const v = cells[k]?.[c];
                    if (!v) return <td key={c} className="nil"></td>;
                    if (secondValueKey) {
                      return <td key={c} className={v.v1 ? 'has-v' : 'nil'}>{v.v1 ? `${v.v1} / ${v.v2}` : ''}</td>;
                    }
                    return <td key={c} className={v.v1 ? 'has-v' : 'nil'}>{v.v1 || ''}</td>;
                  })}
                  <td className="t-col">{secondValueKey ? `${rowTotal1} / ${rowTotal2}` : rowTotal1}</td>
                </tr>
              );
            })}
            <tr className="tot-row">
              <td className="s-col">Σ Total</td>
              {cols.map(c => {
                if (secondValueKey) return <td key={c}>{colTotals[c]?.v1 || 0} / {colTotals[c]?.v2 || 0}</td>;
                return <td key={c}>{colTotals[c]?.v1 || 0}</td>;
              })}
              <td className="t-col">
                {(() => {
                  const gt1 = filtered.reduce((s, r) => s + (r[valueKey] || 0), 0);
                  if (secondValueKey) {
                    const gt2 = filtered.reduce((s, r) => s + (r[secondValueKey] || 0), 0);
                    return `${gt1} / ${gt2}`;
                  }
                  return gt1;
                })()}
              </td>
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
  const lang = 'en';
  const [period, setPeriod] = useState('');
  const [availablePeriods, setAvailablePeriods] = useState([]);
  const [team, setTeam] = useState('all');
  const [shift, setShift] = useState('all'); // Default is Both
  const [timeGrain, setTimeGrain] = useState('all'); // 'all' | 'biweekly1' | 'biweekly2' | 'week1' | 'week2' | 'week3' | 'week4' | 'daily'
  const [selectedDate, setSelectedDate] = useState('');
  const [search, setSearch] = useState('');
  const [specPivotMode, setSpecPivotMode] = useState('class'); // 'class' | 'user'
  const [prodPivotMode, setProdPivotMode] = useState('spec');  // 'spec' | 'user'

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

  // Day-of-month ranges for each slice. Single source of truth — the server-side
  // range query and the client-side filter both read from here so they cannot drift.
  // (Previously there was also a getTimeGrainRatio() that scaled month totals by a
  //  hardcoded fraction. It estimated rather than filtered, so it has been removed.)
  const GRAIN_RANGES = React.useMemo(() => ({
    biweekly1: [1, 15],
    biweekly2: [16, 31],
    week1: [1, 7],
    week2: [8, 14],
    week3: [15, 21],
    week4: [22, 31],
  }), []);

  // Returns a strict YYYY-MM-DD string, or '' if the input cannot be understood.
  // Never returns a half-parsed value — callers rely on '' meaning "unusable".
  const normalizeDateStr = useCallback((dStr) => {
    if (!dStr) return '';
    // Date objects (e.g. if a driver ever hands back a real timestamp)
    if (dStr instanceof Date) {
      if (isNaN(dStr.getTime())) return '';
      const p = n => String(n).padStart(2, '0');
      return `${dStr.getFullYear()}-${p(dStr.getMonth() + 1)}-${p(dStr.getDate())}`;
    }
    const s = String(dStr).trim();
    if (!s) return '';

    const valid = (y, mo, d) =>
      mo >= 1 && mo <= 12 && d >= 1 && d <= new Date(y, mo, 0).getDate();
    const iso = (y, mo, d) =>
      `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;

    // YYYY-MM-DD — unambiguous, this is what the ingest pipeline produces.
    let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?!\d)/);
    if (m) {
      const [y, mo, d] = [+m[1], +m[2], +m[3]];
      return valid(y, mo, d) ? iso(y, mo, d) : '';
    }

    // DD-MM-YYYY / MM-DD-YYYY. Genuinely ambiguous for days 1-12, so only accept
    // it when one ordering is impossible. A silent wrong guess here used to move
    // rows into the wrong week without any visible symptom.
    m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?!\d)/);
    if (m) {
      const [p1, p2, y] = [+m[1], +m[2], +m[3]];
      const dmy = valid(y, p2, p1);   // p1 = day
      const mdy = valid(y, p1, p2);   // p1 = month
      if (dmy && !mdy) return iso(y, p2, p1);
      if (mdy && !dmy) return iso(y, p1, p2);
      // If ambiguous (both valid), assume DD-MM-YYYY which is standard for most of the world
      if (dmy) return iso(y, p2, p1);
      return '';                      // invalid
    }

    return '';
  }, []);

  // Fails CLOSED: a row that carries no usable date is dropped, not kept.
  // The old version returned true for undated rows, so aggregate tables with no
  // date column (specialty, products) passed through untouched and every slice
  // silently showed full-month totals.
  const filterByTimeGrain = useCallback((rows) => {
    if (!rows || !rows.length || timeGrain === 'all') return rows;
    const range = GRAIN_RANGES[timeGrain];
    const normSelected = timeGrain === 'daily' ? normalizeDateStr(selectedDate) : '';
    if (timeGrain === 'daily' && !normSelected) return [];
    if (timeGrain !== 'daily' && !range) return rows;

    return rows.filter(r => {
      const isoDate = normalizeDateStr(r.visit_date || r.coaching_date || r.date);
      if (!isoDate) return false;
      if (timeGrain === 'daily') return isoDate === normSelected;
      const day = parseInt(isoDate.slice(8, 10), 10);
      return day >= range[0] && day <= range[1];
    });
  }, [timeGrain, selectedDate, normalizeDateStr, GRAIN_RANGES]);

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

  // True when the user asked for a sub-month slice but we have no dated rows to
  // honour it with. Surfaced in the UI instead of being papered over with an estimate.
  const sliceUnavailable = React.useMemo(
    () => timeGrain !== 'all' && !(rawVisits || []).length,
    [timeGrain, rawVisits]
  );

  // Auto-default selectedDate when switching to 'daily' mode if unselected
  useEffect(() => {
    if (timeGrain === 'daily' && !selectedDate) {
      if (rawVisits && rawVisits.length > 0) {
        const sortedDates = rawVisits.map(v => v.visit_date).filter(Boolean).sort();
        if (sortedDates.length > 0) {
          setSelectedDate(sortedDates[sortedDates.length - 1]);
          return;
        }
      }
      if (period) {
        const pDate = new Date(`1 ${period}`);
        if (!isNaN(pDate.getTime())) {
          const y = pDate.getFullYear();
          const m = String(pDate.getMonth() + 1).padStart(2, '0');
          setSelectedDate(`${y}-${m}-15`);
        }
      }
    }
  }, [timeGrain, selectedDate, rawVisits, period]);

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
  const codesKey = visibleCodes ? [...visibleCodes].sort().join(',') : '';
  const currentKey = `${periodLabel}|${codesKey}|${isMgr}|${timeGrain}|${selectedDate || ''}`;

  const load = useCallback(async (force = false) => {
    if (!visibleCodes?.length) { setLoading(false); return; }
    const isAdmin = profile?.role === 'Admin';
    const codes = visibleCodes;
    const cacheKey = `dash_${periodLabel}_${isMgr}_${codesKey}_${timeGrain}_${selectedDate || ''}`;

    // Skip if already fetched this key (tab switch won't retrigger)
    if (!force && fetchedKeyRef.current === currentKey) return;

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
        setSummary(overrideSpecialManagers(parsed.summaries));
        if (parsed.visits) setVisits(parsed.visits);
        fetchedKeyRef.current = currentKey;
        setLoading(false);
        return;
      } catch (e) { /* ignore corrupted cache */ }
    }

    setLoading(true); setError('');
    // Resolve periodLabel + timeGrain (+ selectedDate for daily mode) into
    // real ISO date bounds, so the network fetch itself is scoped to the
    // selected week/day — not just filtered client-side after downloading
    // the whole month every time.
    const periodDate = periodLabel ? new Date(`1 ${periodLabel}`) : null;
    let rangeStart = null, rangeEnd = null;
    if (periodDate && !isNaN(periodDate.getTime())) {
      const y = periodDate.getFullYear();
      const m = periodDate.getMonth();
      const lastDay = new Date(y, m + 1, 0).getDate();
      const pad = (n) => String(n).padStart(2, '0');
      const dayStr = (d) => `${y}-${pad(m + 1)}-${pad(d)}`;

      const grainRange = GRAIN_RANGES[timeGrain];
      if (timeGrain === 'daily' && selectedDate) {
        const norm = normalizeDateStr(selectedDate);
        // An unparseable selected date must not silently widen to the whole month.
        rangeStart = norm || dayStr(1);
        rangeEnd = norm || dayStr(lastDay);
      } else if (grainRange) {
        // Clamp to the real length of this month so week4/biweekly2 don't ask for
        // a 31st that doesn't exist.
        rangeStart = dayStr(Math.min(grainRange[0], lastDay));
        rangeEnd = dayStr(Math.min(grainRange[1], lastDay));
      } else {
        // 'all' (whole month) — still bounded, just the full month
        rangeStart = dayStr(1); rangeEnd = dayStr(lastDay);
      }
    }

    const [teamsRes, visitsRes] = await Promise.all([
      supabase.from('teams').select('id, name'),
      (() => {
        let visitsQuery = supabase.from('visits')
          .select('user,employee_code,visit_date,visit_time,shift,acc_type_category,acc_type_raw,visit_type_category,doctor_name,doctor_key,acc_name,acc_id,team,specialty,classification,products')
          .in('employee_code', codes);
        if (rangeStart && rangeEnd) {
          visitsQuery = visitsQuery.gte('visit_date', rangeStart).lte('visit_date', rangeEnd);
        } else if (periodLabel) {
          visitsQuery = visitsQuery.eq('period', periodLabel);
        }
        return visitsQuery;
      })()
    ]);

    if (visitsRes.data) setVisits(visitsRes.data);

    let tMap = {};
    if (teamsRes.data) {
      teamsRes.data.forEach(t => tMap[t.id] = t.name);
      setTeamsMap(tMap);
    }

    if (visitsRes.error) {
      setError(visitsRes.error.message);
    } else {
      // Build base summary from hierarchy so we have a row for every visible user
      const baseSummaries = (hierarchy || [])
        .filter(h => codes.includes(h.employee_code))
        .map(h => ({
          user_name: h.employee_name,
          employee_code: h.employee_code,
          role: h.role,
          is_manager: h.role !== 'MR',
          team: tMap[h.team_id] || 'Unknown',
          territory: h.division_name || '',
          manager_name: h.supervisor_name || h.area_manager_name || h.blm_name || '',
        }));
      
      const dataToCache = { summaries: baseSummaries, visits: visitsRes.data };
      try {
        sessionStorage.setItem(cacheKey, JSON.stringify(dataToCache));
      } catch (e) { }

      setSummary(overrideSpecialManagers(baseSummaries));
    }

    setSpecialty([]);
    setProducts([]);
    setCoaching([]);
    fetchedKeyRef.current = currentKey;
    setLoading(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [periodLabel, currentKey, isMgr, profile, timeGrain, selectedDate]);

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
    r.forEach(x => { m.set(x.user_name, { ...x }); });
    const finalArr = Array.from(m.values());

    const formatTimeStr = (totalMins) => {
      if (!totalMins) return '—';
      let hrs = Math.floor(totalMins / 60);
      let mins = Math.floor(totalMins % 60);
      const ampm = hrs >= 12 ? 'PM' : 'AM';
      if (hrs > 12) hrs -= 12;
      if (hrs === 0) hrs = 12;
      return `${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')} ${ampm}`;
    };

    if (rawVisits && rawVisits.length > 0) {
      const processedVisits = filterByTimeGrain(rawVisits).map(v => {
        const cat = (v.acc_type_category || '').toLowerCase();
        let derivedShift = '';
        if (cat.includes('hospital') || cat.includes('am center') || cat.includes('distributer') || cat.includes('distributor')) {
          derivedShift = 'AM';
        } else if (cat.includes('clinic') || cat.includes('poly')) {
          derivedShift = 'PM';
        } else {
          derivedShift = v.shift;
        }

        const isPharmacy = cat.includes('pharmacy') || (v.acc_type_raw || '').toLowerCase().includes('pharmacy') || (v.acc_name || '').toLowerCase().includes('pharmacy');
        const isActivity = cat.includes('activity') || cat.includes('office') || (v.visit_type_category || '').toLowerCase().includes('activity') || (v.visit_type_category || '').toLowerCase().includes('office');

        let timeStr = v.visit_time || '';
        let visitMinutes = 0;
        if (timeStr) {
          if (typeof timeStr === 'number') {
            const totalSeconds = Math.round(timeStr * 24 * 3600);
            visitMinutes = Math.floor(totalSeconds / 60);
          } else {
             const parts = String(timeStr).split(':');
             if (parts.length >= 2) {
               visitMinutes = parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
               if (String(timeStr).toLowerCase().includes('pm') && parseInt(parts[0], 10) !== 12) visitMinutes += 12 * 60;
               else if (String(timeStr).toLowerCase().includes('am') && parseInt(parts[0], 10) === 12) visitMinutes -= 12 * 60;
             }
          }
        }
        return { ...v, derivedShift, isPharmacy, isActivity, visitMinutes };
      });

      const doubleVisitGroups = new Map();
      processedVisits.forEach(v => {
        if (!v.acc_name || !v.visit_date || !v.visit_time || !v.specialty || v.isPharmacy || v.isActivity) return;
        const key = `${v.acc_name}||${v.visit_date}||${v.visit_time}||${v.specialty}`;
        if (!doubleVisitGroups.has(key)) doubleVisitGroups.set(key, new Set());
        doubleVisitGroups.get(key).add(v.user?.toLowerCase().trim());
      });

      finalArr.forEach(x => {
        const normName = x.user_name?.toLowerCase().trim();
        const code = x.employee_code;
        const userVisits = processedVisits.filter(v => (code && String(v.employee_code).trim() === String(code).trim()) || (normName && v.user?.toLowerCase().trim() === normName));

        // Find days with PM Activity or PM Office Work
        const pmActivityDays = new Set(userVisits.filter(v => v.isActivity && (v.shift === 'PM' || v.derivedShift === 'PM' || v.visitMinutes >= 12 * 60)).map(v => v.visit_date).filter(Boolean));

        // Filter out visits that land on a PM activity day (except for total raw visits which counts them)
        const validVisits = userVisits.filter(v => !pmActivityDays.has(v.visit_date));

        const validVisitsByDate = new Map();
        validVisits.forEach(v => {
           if (!v.visit_date) return;
           if (!validVisitsByDate.has(v.visit_date)) validVisitsByDate.set(v.visit_date, []);
           validVisitsByDate.get(v.visit_date).push(v);
        });

        let fieldWorkingDaysCount = 0;
        let totalAmFirstVisitMinutes = 0;
        let validAmFirstVisitDays = 0;
        let amShiftDays = 0;
        let pmShiftDays = 0;
        let totalAmShiftDurMinutes = 0;
        let validAmShiftDurDays = 0;
        let totalPmShiftDurMinutes = 0;
        let validPmShiftDurDays = 0;

        validVisitsByDate.forEach((dayVisits) => {
          const doctorVisits = dayVisits.filter(v => !v.isPharmacy && !v.isActivity);
          const hasAm = doctorVisits.some(v => v.derivedShift === 'AM');
          const hasPm = doctorVisits.some(v => v.derivedShift === 'PM');
          
          if (hasAm || hasPm) {
            fieldWorkingDaysCount++;
            if (hasAm) amShiftDays++;
            if (hasPm) pmShiftDays++;

            const amDoctorVisits = doctorVisits.filter(v => v.derivedShift === 'AM' && v.visitMinutes > 0);
            if (amDoctorVisits.length > 0) {
              const times = amDoctorVisits.map(v => v.visitMinutes).sort((a,b)=>a-b);
              totalAmFirstVisitMinutes += times[0];
              validAmFirstVisitDays++;
              if (times.length > 1) {
                totalAmShiftDurMinutes += (times[times.length - 1] - times[0]);
                validAmShiftDurDays++;
              }
            }

            const pmDoctorVisits = doctorVisits.filter(v => v.derivedShift === 'PM' && v.visitMinutes > 0);
            if (pmDoctorVisits.length > 1) {
              const times = pmDoctorVisits.map(v => v.visitMinutes).sort((a,b)=>a-b);
              totalPmShiftDurMinutes += (times[times.length - 1] - times[0]);
              validPmShiftDurDays++;
            }
          }
        });

        const amDoctorVisits = validVisits.filter(v => v.derivedShift === 'AM' && !v.isPharmacy && !v.isActivity);
        x.am_calls = amDoctorVisits.length;
        x.amcenter_covered = new Set(amDoctorVisits.map(v => v.doctor_key || v.doctor_name).filter(Boolean)).size;
        x.am_accounts_unique = new Set(amDoctorVisits.map(v => v.acc_id || v.acc_name).filter(Boolean)).size;

        const pmDoctorVisits = validVisits.filter(v => v.derivedShift === 'PM' && !v.isPharmacy && !v.isActivity);
        x.pm_calls = pmDoctorVisits.length;
        x.clinic_covered = new Set(pmDoctorVisits.map(v => v.doctor_key || v.doctor_name).filter(Boolean)).size;
        x.polyclinic_covered = new Set(pmDoctorVisits.map(v => v.acc_id || v.acc_name).filter(Boolean)).size;

        const pharmacyVisits = validVisits.filter(v => v.isPharmacy);
        x.pharmacies_visited = pharmacyVisits.length;
        x.pharmacies_covered = new Set(pharmacyVisits.map(v => v.acc_id || v.acc_name).filter(Boolean)).size;

        let doubleVisitsCount = 0;
        validVisits.forEach(v => {
          if (!v.acc_name || !v.visit_date || !v.visit_time || !v.specialty || v.isPharmacy || v.isActivity) return;
          const key = `${v.acc_name}||${v.visit_date}||${v.visit_time}||${v.specialty}`;
          const usersInVisit = doubleVisitGroups.get(key);
          if (usersInVisit && usersInVisit.size > 1 && usersInVisit.has(normName)) {
            doubleVisitsCount++;
          }
        });
        x.double_visit_days = doubleVisitsCount;

        x.working_days = fieldWorkingDaysCount;
        x.am_shift_days = amShiftDays;
        x.pm_shift_days = pmShiftDays;

        x.avg_am_start_time = validAmFirstVisitDays > 0 ? formatTimeStr(totalAmFirstVisitMinutes / validAmFirstVisitDays) : '—';
        x.avg_am_shift_hm = validAmShiftDurDays > 0 ? (totalAmShiftDurMinutes / validAmShiftDurDays) / 60 : 0;
        x.avg_pm_shift_hm = validPmShiftDurDays > 0 ? (totalPmShiftDurMinutes / validPmShiftDurDays) / 60 : 0;

        x.total_visits = userVisits.length;
        x.no_activities = 0;
        x.no_events = 0;
        x.office_work_days = pmActivityDays.size;
        x.total_am_covered = x.amcenter_covered;
        x.total_pm_covered = x.clinic_covered + x.polyclinic_covered;
        x.am_call_rate = amShiftDays > 0 ? Math.round((x.am_calls / amShiftDays) * 10) / 10 : 0;
        x.pm_call_rate = pmShiftDays > 0 ? Math.round((x.pm_calls / pmShiftDays) * 10) / 10 : 0;
      });
    }

    return sortSummary(finalArr);
  }, [summary, rawVisits, byTeam, byLineManager, byManagerTerritory, search, userFilter, hierarchy, filterByTimeGrain]);

  const managerNames = useMemo(() => {
    const s = new Set();
    summary.forEach(r => {
      if (r.is_manager) s.add((r.user_name || '').toLowerCase().trim());
    });
    return s;
  }, [summary]);

  const fSpecialty = useMemo(() => {
    if (rawVisits && rawVisits.length > 0 && rawVisits.some(v => v.specialty)) {
      const filteredVisits = filterByTimeGrain(rawVisits);
      const map = new Map();
      filteredVisits.forEach(v => {
        if (!v.specialty) return;
        const u = v.user || v.user_name || 'Unknown';
        const spec = v.specialty;
        const cls = v.classification || 'Unclassified';
        
        const cat = (v.acc_type_category || '').toLowerCase();
        let shft = '';
        if (cat.includes('hospital') || cat.includes('am center') || cat.includes('distributer') || cat.includes('distributor')) {
          shft = 'AM';
        } else if (cat.includes('clinic') || cat.includes('poly')) {
          shft = 'PM';
        } else {
          shft = v.shift || 'AM';
        }

        const isPharmacy = cat.includes('pharmacy') || (v.acc_type_raw || '').toLowerCase().includes('pharmacy') || (v.acc_name || '').toLowerCase().includes('pharmacy');
        const isActivity = cat.includes('activity') || cat.includes('office') || (v.visit_type_category || '').toLowerCase().includes('activity') || (v.visit_type_category || '').toLowerCase().includes('office');
        
        if (isPharmacy || isActivity) return;

        const key = `${u}||${spec}||${cls}||${shft}`;
        if (!map.has(key)) {
          map.set(key, {
            user_name: u,
            employee_code: v.employee_code,
            specialty: spec,
            classification: cls,
            shift: shft,
            call_count: 0,
            _coveredSet: new Set(),
            covered: 0,
            team: userTeamMap[u.toLowerCase().trim()] || v.team || ''
          });
        }
        
        const item = map.get(key);
        item.call_count += 1;
        if (v.doctor_key || v.doctor_name) {
           item._coveredSet.add(v.doctor_key || v.doctor_name);
        }
      });
      let r = Array.from(map.values()).map(x => { x.covered = x._coveredSet.size; return x; });
      r = byManagerTerritory(byLineManager(byTeam(r)));
      if (search) r = r.filter(x => x.user_name?.toLowerCase().includes(search.toLowerCase()) || x.territory?.toLowerCase().includes(search.toLowerCase()));
      if (userFilter !== 'all') r = r.filter(x => x.user_name === userFilter);
      if (userFilter === 'all') r = r.filter(x => !managerNames.has((x.user_name || '').toLowerCase().trim()));
      return r;
    }
    return [];
  }, [rawVisits, filterByTimeGrain, byTeam, byLineManager, byManagerTerritory, search, userFilter, managerNames, userTeamMap]);

  const fProducts = useMemo(() => {
    if (rawVisits && rawVisits.length > 0 && rawVisits.some(v => v.products)) {
      const filteredVisits = filterByTimeGrain(rawVisits);
      const map = new Map();
      filteredVisits.forEach(v => {
        if (!v.products) return;
        const u = v.user || v.user_name || 'Unknown';
        const spec = v.specialty || 'General';
        
        const cat = (v.acc_type_category || '').toLowerCase();
        let shft = '';
        if (cat.includes('hospital') || cat.includes('am center') || cat.includes('distributer') || cat.includes('distributor')) {
          shft = 'AM';
        } else if (cat.includes('clinic') || cat.includes('poly')) {
          shft = 'PM';
        } else {
          shft = v.shift || 'AM';
        }

        const isPharmacy = cat.includes('pharmacy') || (v.acc_type_raw || '').toLowerCase().includes('pharmacy') || (v.acc_name || '').toLowerCase().includes('pharmacy');
        const isActivity = cat.includes('activity') || cat.includes('office') || (v.visit_type_category || '').toLowerCase().includes('activity') || (v.visit_type_category || '').toLowerCase().includes('office');
        
        if (isPharmacy || isActivity) return;

        const prods = v.products.split(',').map(p => p.trim()).filter(Boolean);
        prods.forEach(prod => {
          const key = `${u}||${spec}||${prod}||${shft}`;
          if (!map.has(key)) {
            map.set(key, {
              user_name: u,
              employee_code: v.employee_code,
              specialty: spec,
              product: prod,
              shift: shft,
              call_count: 0,
              _coveredSet: new Set(),
              covered: 0,
              team: userTeamMap[u.toLowerCase().trim()] || v.team || ''
            });
          }
          const item = map.get(key);
          item.call_count += 1;
          if (v.doctor_key || v.doctor_name) {
             item._coveredSet.add(v.doctor_key || v.doctor_name);
          }
        });
      });
      let r = Array.from(map.values()).map(x => { x.covered = x._coveredSet.size; return x; });
      r = byManagerTerritory(byLineManager(byTeam(r)));
      if (search) r = r.filter(x => x.user_name?.toLowerCase().includes(search.toLowerCase()) || x.territory?.toLowerCase().includes(search.toLowerCase()));
      if (userFilter !== 'all') r = r.filter(x => x.user_name === userFilter);
      if (userFilter === 'all') r = r.filter(x => !managerNames.has((x.user_name || '').toLowerCase().trim()));
      return r;
    }
    return [];
  }, [rawVisits, filterByTimeGrain, byTeam, byLineManager, byManagerTerritory, search, userFilter, managerNames, userTeamMap]);

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
    let r = fSummary.filter(x => x.is_manager && (x.role === 'Area Manager' || x.role === 'Supervisor'));
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

    return r.map(x => {
       const target = x.role === 'Area Manager' ? 6 : (x.role === 'Supervisor' ? 10 : 0);
       const achieved = x.double_visit_days || 0;
       const pct = target > 0 ? (achieved / target) * 100 : (achieved > 0 ? 100 : 0);
       return { ...x, target, achieved, pct };
    });
  }, [fSummary, search, userFilter, hierarchy]);

  // ── Timing data: last visit time per rep per day ──────────────────────
  const timingData = useMemo(() => {
    if (!rawVisits?.length) return [];

    const processedVisits = filterByTimeGrain(rawVisits).map(v => {
      const cat = (v.acc_type_category || '').toLowerCase();
      let derivedShift = '';
      if (cat.includes('hospital') || cat.includes('am center') || cat.includes('distributer') || cat.includes('distributor')) derivedShift = 'AM';
      else if (cat.includes('clinic') || cat.includes('poly')) derivedShift = 'PM';
      else derivedShift = v.shift;

      const isPharmacy = cat.includes('pharmacy') || (v.acc_type_raw || '').toLowerCase().includes('pharmacy') || (v.acc_name || '').toLowerCase().includes('pharmacy');
      const isActivity = cat.includes('activity') || cat.includes('office') || (v.visit_type_category || '').toLowerCase().includes('activity') || (v.visit_type_category || '').toLowerCase().includes('office');
      
      let visitMinutes = 0;
      let timeStr = v.visit_time || '';
      if (timeStr) {
        if (typeof timeStr === 'number') {
          visitMinutes = Math.floor(Math.round(timeStr * 24 * 3600) / 60);
        } else {
           const parts = String(timeStr).split(':');
           if (parts.length >= 2) {
             visitMinutes = parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
             if (String(timeStr).toLowerCase().includes('pm') && parseInt(parts[0], 10) !== 12) visitMinutes += 12 * 60;
             else if (String(timeStr).toLowerCase().includes('am') && parseInt(parts[0], 10) === 12) visitMinutes -= 12 * 60;
           }
        }
      }
      return { ...v, derivedShift, isPharmacy, isActivity, visitMinutes };
    });

    const pmActivityDates = new Set();
    processedVisits.forEach(v => {
      if (v.isActivity && (v.shift === 'PM' || v.derivedShift === 'PM' || v.visitMinutes >= 12 * 60) && v.user && v.visit_date) {
        pmActivityDates.add(`${v.user}|||${v.visit_date}`);
      }
    });

    // Include only valid PM visits for timing data
    const validVisits = processedVisits.filter(v => v.visit_date && v.visitMinutes > 0 && v.derivedShift === 'PM' && !v.isPharmacy && !v.isActivity);

    const byUserDate = {};
    validVisits.forEach(v => {
      const user = v.user || '';
      const date = v.visit_date || '';
      if (!user || !date) return;
      const key = `${user}|||${date}`;

      if (pmActivityDates.has(key)) return;

      if (!byUserDate[key] || v.visitMinutes > byUserDate[key].visitMinutes) {
        byUserDate[key] = { user, date, time: v.visit_time, visitMinutes: v.visitMinutes, team: v.team || '', employee_code: v.employee_code };
      }
    });

    const categorizeTime = (mins) => {
      if (!mins) return 'unknown';
      if (mins < 15 * 60) return 'early';
      if (mins <= 18 * 60) return 'normal';
      return 'late';
    };

    const formatTimeStr = (totalMins) => {
      if (!totalMins) return '—';
      let hrs = Math.floor(totalMins / 60);
      let mins = Math.floor(totalMins % 60);
      const ampm = hrs >= 12 ? 'PM' : 'AM';
      if (hrs > 12) hrs -= 12;
      if (hrs === 0) hrs = 12;
      return `${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')} ${ampm}`;
    };

    return Object.values(byUserDate).map(entry => ({
      ...entry,
      category: categorizeTime(entry.visitMinutes),
      formattedTime: formatTimeStr(entry.visitMinutes),
    }));
  }, [rawVisits, filterByTimeGrain]);

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

  // Timing category filter state
  const [timingCategoryFilter, setTimingCategoryFilter] = useState('all');

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
          r.date ? formatDateDisplay(r.date) : '—',
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
          r.manager_name || '—', r.rep_name || '—', r.coaching_date ? formatDateDisplay(r.coaching_date) : '—', r.team || '—',
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
                    <option key={p} value={p}>{formatDateDisplay(p)}</option>
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
                  <option value="daily">{rtl ? 'يوم محدد' : 'Specific Day'}</option>
                </select>
              </div>
              {timeGrain === 'daily' && (
                <div className="ctrl-group">
                  <span className="ctrl-lbl">{rtl ? 'التاريخ' : 'Date'}</span>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <input
                      type="date"
                      className="ctrl-sel"
                      value={selectedDate}
                      onChange={e => setSelectedDate(e.target.value)}
                    />
                    {selectedDate && (
                      <span className="kpi-tag" style={{ whiteSpace: 'nowrap', fontSize: '12px', background: '#3b82f6', color: '#fff', padding: '4px 8px', borderRadius: '4px' }}>
                        📅 {formatDateDisplay(selectedDate)}
                      </span>
                    )}
                  </div>
                </div>
              )}
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

              {sliceUnavailable && (
                <div className="dash-slice-warning" role="status">{t.sliceUnavailable}</div>
              )}
              {timeGrain !== 'all' && (tab === 'specialty' || tab === 'products') && !(rawVisits || []).some(v => tab === 'specialty' ? v.specialty : v.products) && (
                <div className="dash-slice-warning" role="status">{t.sliceNoDates}</div>
              )}

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
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                    <div style={{ flex: 1 }}>
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
                    </div>
                    <div className="shift-toggle" style={{ marginLeft: '12px', flexShrink: 0 }}>
                      <button className={`stoggle${specPivotMode === 'class' ? ' on' : ''}`} onClick={() => setSpecPivotMode('class')}>
                        📊 {rtl ? 'التخصص × التصنيف' : 'Specialty × Classification'}
                      </button>
                      <button className={`stoggle${specPivotMode === 'user' ? ' on' : ''}`} onClick={() => setSpecPivotMode('user')}>
                        👤 {rtl ? 'التخصص × المندوب' : 'Specialty × Rep'}
                      </button>
                    </div>
                  </div>
                  <PivotTable
                    rows={fSpecialty}
                    rowKey="specialty"
                    colKey={specPivotMode === 'class' ? 'classification' : 'user_name'}
                    valueKey="call_count"
                    secondValueKey="covered"
                    shiftFilter={shift}
                    userFilter={userFilter}
                    searchFilter={search}
                    lang={lang}
                    hideAvg={specPivotMode === 'class'}
                    colTitle={specPivotMode === 'class' ? (rtl ? 'التصنيف' : 'Classification') : (rtl ? 'المندوب' : 'Rep')}
                  />
                </>
              )}

              {/* PRODUCTS TAB */}
              {tab === 'products' && (
                <>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                    <div style={{ flex: 1 }}>
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
                    </div>
                    <div className="shift-toggle" style={{ marginLeft: '12px', flexShrink: 0 }}>
                      <button className={`stoggle${prodPivotMode === 'spec' ? ' on' : ''}`} onClick={() => setProdPivotMode('spec')}>
                        💊 {rtl ? 'المنتج × التخصص' : 'Product × Specialty'}
                      </button>
                      <button className={`stoggle${prodPivotMode === 'user' ? ' on' : ''}`} onClick={() => setProdPivotMode('user')}>
                        👤 {rtl ? 'المنتج × المندوب' : 'Product × Rep'}
                      </button>
                    </div>
                  </div>
                  <PivotTable
                    rows={fProducts}
                    rowKey="product"
                    colKey={prodPivotMode === 'spec' ? 'specialty' : 'user_name'}
                    valueKey="call_count"
                    secondValueKey="covered"
                    shiftFilter={shift}
                    userFilter={userFilter}
                    searchFilter={search}
                    lang={lang}
                    hideAvg={false}
                    colTitle={prodPivotMode === 'spec' ? (rtl ? 'التخصص' : 'Specialty') : (rtl ? 'المندوب' : 'Rep')}
                  />
                </>
              )}

              {/* COACHING TAB */}
              {tab === 'coaching' && isMgr && (
                fCoaching.length === 0 ? (
                  <div className="dash-empty">
                    {t.noData}
                  </div>
                ) : (
                  <>
                    <div className="pivot-wrap">
                      <table className="pivot-tbl">
                        <thead>
                          <tr>
                            <th className="s-col">{rtl ? 'المدير' : 'Manager'}</th>
                            <th>{rtl ? 'الدور' : 'Role'}</th>
                            <th>{rtl ? 'المنطقة' : 'Territory'}</th>
                            <th>{rtl ? 'أيام المرافقة الفعلية' : 'Achieved Coaching Days'}</th>
                            <th>{rtl ? 'الهدف' : 'Target'}</th>
                            <th>% {rtl ? 'من الهدف' : 'of Target'}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {[...fCoaching].sort((a, b) => b.pct - a.pct).map((r, i) => (
                            <tr key={r.employee_code || i}>
                              <td className="s-col">{r.user_name}</td>
                              <td>{r.role}</td>
                              <td>{r.territory}</td>
                              <td>{r.achieved}</td>
                              <td>{r.target}</td>
                              <td style={{ color: r.pct < 80 ? '#ef4444' : 'inherit', fontWeight: r.pct < 80 ? '600' : 'normal' }}>
                                {Math.round(r.pct)}%
                              </td>
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
                          {filteredTiming.map((r, i) => (
                            <tr key={`${r.user}-${r.date}-${i}`} className={`timing-row timing-row-${r.category}`}>
                              <td className="s-col">{formatDateDisplay(r.date)}</td>
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