# Refactor Dashboard to use Raw Visits (Full System Update)

We will completely rewrite the data calculation logic in `Dashboard.js` to process **all** reports directly from the raw `visits` data. This ensures absolute accuracy and implements all of your specific business rules.

## Global Rules & Definitions

> [!IMPORTANT]
> **PM Activity / Office Work Override:**
> If a user has a PM Activity or PM Office Work on a specific date, **that entire day will be excluded** from ALL metrics (Working Days, Shift Durations, AM/PM calls, Unique covers, Timing stats). The visits from that day will ONLY be counted in a basic "Total Number Visited" metric, as requested.

1. **AM Shift Definition:** Accounts that are AM Center, Hospitals, or Distributors.
2. **PM Shift Definition:** Accounts that are Clinics or Polyclinics.
3. **Pharmacies:** Will be strictly excluded from ANY shift calculations (like AM/PM durations or AM/PM counts). They will only be counted in their own separate Pharmacy metrics.

---

## Proposed Changes by Report

### 1. Summary Report
We will calculate the following exactly as requested per user:
- **Field Working Days:** Count of days where there is at least one AM or PM shift visit to a doctor (excluding pharmacies, activities, and office work, and excluding days with PM activity).
- **AM Metrics:** 
  - Total visits to AM Doctors
  - Unique AM Doctors visited
  - Unique AM Accounts visited
- **PM Metrics:** 
  - Total visits to PM Doctors
  - Unique PM Doctors visited
  - Unique PM Accounts visited
- **Pharmacy Metrics:**
  - Total Pharmacy visits
  - Unique Pharmacies visited
- **Double Visits:** Calculated by scanning all visits across the team. Any visits that share the identical `acc_name`, `date`, `time`, and `specialty` (but a different user) will add +1 to the user's Double Visits count.
- **Timing:** 
  - **Avg AM First Visit:** Average of the earliest AM doctor visit time across valid working days.
  - **Avg Shift Duration (AM/PM):** Average of (Last Visit Time - First Visit Time) per shift across valid working days, strictly excluding Pharmacies.

### 2. Specialty Report
- We will group the data by **Specialty**, **Classification**, and **Shift (AM/PM)**.
- For each grouping, we will calculate and display:
  - **Visited:** Total number of visits.
  - **Covered:** Unique doctors covered.

### 3. Products Report
- We will group the data by **Product**, **Team User**, and **Shift (AM/PM)**.
- Similar to Specialty, we will display:
  - **Calls:** Total number of calls for that product.
  - **Unique Covers:** Unique doctors detailed on that product.

### 4. Coaching / Double Visits (DV) for Managers
- We will group double visits by Manager and their Team Members.
- We will display:
  - Number of double visits (and how many were AM vs PM).
  - Frequency of double visits per team member.
  - **Highlighting Rule:** We will compare the Manager's Double Visits with a Rep during a shift against the Rep's Total Visits for that shift. If the Manager did not cover at least **80%** of the Rep's visits for that shift on that day, it will be visually highlighted (e.g., in red or with a warning icon).

### 5. Last Visit Time (The Day)
- We will identify the time of the **Last PM Shift Visit** for each user per day (excluding days with PM Activity/Office Work).
- We will segment these times into:
  - **Before 3 PM:** Bad (Visually flagged)
  - **3 PM - 6 PM:** Moderate
  - **After 6 PM:** Good
- This section will display these segmented counts for the selected period.

## Open Questions

> [!WARNING]
> Please review and confirm the following:
> 1. **AM/PM Definitions:** You defined AM shift as "AM center, hospitals, and distributors" and PM shift as "clinic/polyclinic". Currently, the raw excel data has a `shift` column (which says "AM" or "PM"). Should I rely on the Excel `shift` column, or should I forcefully recategorize the shift based on the account type (e.g. if the Excel says 'PM' but it's a Hospital, should I force it to be 'AM')?
> 2. **Double Visits Calculation:** For the Summary tab "Double Visits" count (where two reps have identical visit details), should this look at ALL visits (including Pharmacies), or strictly Doctor visits?
