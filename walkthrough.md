# Dashboard Raw Data Refactoring Walkthrough

The application's `Dashboard.js` component has been entirely refactored to compute all reports strictly from the raw `visits` data fetched from Supabase, completely bypassing the legacy `get_dashboard_data` RPC. This ensures 100% accuracy based directly on the raw Excel-like records, eliminating any residual or mismatched aggregations.

## Architectural Changes

1. **Removed RPC Dependency**
   - The `supabase.rpc('get_dashboard_data', ...)` call was removed from the `load()` function.
   - The application now solely queries `supabase.from('visits').select(...)` and `supabase.from('teams').select(...)`.
   - The initial `summary` state is now populated using the current user's visible `hierarchy`, ensuring every rep has a baseline record, even if they have no visits.

2. **Unified Data Processing**
   - All logic previously handled by the backend has been migrated directly into frontend `useMemo` hooks (`fSummary`, `fSpecialty`, `fProducts`, `fCoaching`, `fTiming`).
   - Time calculations, AM/PM shifting, and double visit tracking are derived directly from the raw records on the client side.

## Detailed Report Logic Implementations

### Summary Tab
- **Field Working Days**: Exclusively counts days where at least one valid AM or PM doctor visit occurred. Any day containing a PM Activity or Office Work is entirely excluded from working days.
- **AM/PM Metrics**: 
  - Doctor visits are correctly segmented into AM (Hospital, AM Center, Distributor) and PM (Clinic, Polyclinic).
  - Pharmacies and Activities are excluded from the core Shift Doctor KPIs.
- **Pharmacy Metrics**: Pharmacies are now tracked as a separate, distinct KPI group (Pharmacies Visited, Pharmacies Covered) and do not inflate the AM/PM call metrics.
- **Double Visits**: Computed by scanning all records and finding instances where multiple unique reps have an exact match on `acc_name`, `visit_date`, `visit_time`, and `specialty`.
- **Timing**: Average AM Start Time, Average AM Duration, and Average PM Duration are now precisely derived from the actual visit timestamps.

### Specialty & Products Tabs
- The pivot tables now display both **Visits (Calls)** and **Unique Coverage (Doctors Covered)** per classification or product. 
- Pharmacies and Activities are excluded from these specific medical/product aggregates to keep the data clean.
- Display logic in `PivotTable` was updated to support the `calls / covered` display format.

### Coaching Tab
- The Coaching tab has been completely redesigned to track **Manager Double Visits**.
- It now displays the Area Manager and Supervisor Target Coaching Days (6 for AMs, 10 for Supervisors) versus what they actually achieved on the field.
- Managers falling below 80% of their target will have their percentage highlighted in red.

### Last Visit Data (Timing Tab)
- Accurately captures the final valid PM visit of the day for each user, sorting them into Early (Before 3 PM), Normal (3 PM - 6 PM), and Late (After 6 PM).
- Days with PM Activity or Office Work are fully ignored for these users, avoiding false positives on early finish times.

## Next Steps
Please thoroughly review the Dashboard using your staging or production data. Ensure the numbers align perfectly with the raw data source and that the dynamic sub-month slicing now works cleanly without the previous RPC limitations.
