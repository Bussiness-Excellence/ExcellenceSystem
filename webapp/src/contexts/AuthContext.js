import React, { createContext, useContext, useEffect, useState, useCallback } from 'react';
import { supabase, employeeCodeToEmail } from '../supabaseClient';

const AuthContext = createContext(null);

export function useAuth() {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used inside <AuthProvider>');
  return ctx;
}

/**
 * Compute which employee_codes this user can see, based on the hierarchy chain.
 *
 * How the hierarchy table works:
 *   - Every row has: employee_name, employee_code, role, supervisor_name, team_id
 *   - MRs:           supervisor_name = their supervisor's name
 *   - Supervisors:   supervisor_name = their own name (self-referential)
 *   - Area Managers: supervisor_name = their own name
 *   - BLMs:          supervisor_name = null
 *
 * Visibility rules:
 *   MR           → own code only
 *   Supervisor   → own code + codes of all MRs whose supervisor_name = my name
 *   Area Manager → own code + codes of supervisors in my team +
 *                  codes of all MRs under those supervisors
 *                  BUT only supervisors that appear under ME in the hierarchy,
 *                  detected by: MRs whose supervisor_name matches a supervisor
 *                  who is in the same team and whose area_manager_name = my name
 *                  (fallback: ALL supervisors in team if area_manager_name not set)
 *   BLM          → everyone in my team
 *   Admin        → everyone
 */
export function computeVisibleEmployeeCodes(profile, hierarchyRows, teamsRows) {
  if (!profile) return [];
  const { role, employee_name: myName, team_id: myTeamId, employee_code: myCode } = profile;

  if (role === 'Admin') {
    return hierarchyRows.map(h => h.employee_code).filter(Boolean);
  }

  if (role === 'Stakeholder') {
    const visibleTeamNames = profile.visible_teams || [];
    if (!visibleTeamNames.length) {
      return hierarchyRows.map(h => h.employee_code).filter(Boolean);
    }
    const teamNameMap = {};
    (teamsRows || []).forEach(t => teamNameMap[t.name] = t.id);
    const visibleTeamIds = new Set(visibleTeamNames.map(n => teamNameMap[n]));
    
    const visible = hierarchyRows
      .filter(h => visibleTeamIds.has(h.team_id))
      .map(h => h.employee_code);
    return [...new Set([myCode, ...visible])].filter(Boolean);
  }

  if (role === 'BLM') {
    const myTeamIds = new Set(
      hierarchyRows
        .filter(h => h.employee_name === myName && h.role === 'BLM')
        .map(h => h.team_id)
    );
    if (myTeamIds.size === 0 && myTeamId) {
      myTeamIds.add(myTeamId);
    }
    return hierarchyRows
      .filter(h => myTeamIds.has(h.team_id))
      .map(h => h.employee_code)
      .filter(Boolean);
  }

  if (role === 'Area Manager') {
    const myTeamIds = new Set(
      hierarchyRows
        .filter(h => h.employee_name === myName && h.role === 'Area Manager')
        .map(h => h.team_id)
    );
    if (myTeamIds.size === 0 && myTeamId) {
      myTeamIds.add(myTeamId);
    }

    const myTeamRows = hierarchyRows.filter(h => myTeamIds.has(h.team_id));

    // Supervisors reporting to me
    const mySupervisorNames = new Set(
      myTeamRows
        .filter(h => (h.area_manager_name === myName || h.supervisor_name === myName) && h.role === 'Supervisor')
        .map(h => h.employee_name)
    );

    // Include self, direct reports (area_manager_name/supervisor_name === myName), supervisors, and MRs under my supervisors
    const visible = myTeamRows
      .filter(h =>
        h.employee_name === myName ||
        h.area_manager_name === myName ||
        h.supervisor_name === myName ||
        mySupervisorNames.has(h.employee_name) ||
        mySupervisorNames.has(h.supervisor_name)
      )
      .map(h => h.employee_code);

    return [...new Set([myCode, ...visible])].filter(Boolean);
  }

  if (role === 'Supervisor') {
    // Include self + all MRs reporting to this supervisor
    const visible = hierarchyRows
      .filter(h =>
        h.employee_name === myName ||
        h.supervisor_name === myName ||
        h.area_manager_name === myName
      )
      .map(h => h.employee_code);
    return [...new Set([myCode, ...visible])].filter(Boolean);
  }

  // MR default
  return [myCode].filter(Boolean);
}

export function AuthProvider({ children }) {
  const [session, setSession]         = useState(null);
  const [profile, setProfile]         = useState(null);
  const [hierarchy, setHierarchy]     = useState([]);
  const [visibleCodes, setVisibleCodes] = useState([]);
  const [loading, setLoading]         = useState(true);
  // Tracks which user's profile/hierarchy is currently loaded, so a
  // redundant SIGNED_IN event (e.g. fired when the browser tab regains
  // focus) doesn't re-trigger a full loading state for a user we already have.
  const loadedUserId = React.useRef(null);

  const loadProfileAndHierarchy = useCallback(async (userId) => {
    let [{ data: profileRow }, { data: hierarchyRows }, { data: teamsRows }] =
      await Promise.all([
        supabase.from('app_users').select('*').eq('id', userId).maybeSingle(),
        supabase.from('hierarchy').select('*').range(0, 5000),
        supabase.from('teams').select('*'),
      ]);

    if (!profileRow) {
      try {
        const { data: authUserData } = await supabase.auth.getUser();
        const authUser = authUserData?.user;
        if (authUser?.email) {
          const empCode = authUser.email.split('@')[0];
          const { data: matchedProfile } = await supabase.from('app_users')
            .select('*')
            .or(`employee_code.eq.${empCode},email.eq.${authUser.email}`)
            .maybeSingle();

          if (matchedProfile) {
            profileRow = matchedProfile;
            supabase.from('app_users').update({ id: userId }).eq('employee_code', matchedProfile.employee_code).then(() => {});
          }
        }
      } catch (e) {
        console.error('Profile fallback lookup error:', e);
      }
    }

    if (!profileRow) {
      console.error('Profile not found for user ID:', userId);
      setProfile(null); setVisibleCodes([]);
      return;
    }

    const SPECIAL_MANAGERS = [
      'ahmad morsy', 'ahmed elasyed', 'ahmed tarek mohamed', 'akram ahmed elhossary',
      'asmaa abdel fattah', 'dm', 'evette zakaria hefni', 'gihad sayed', 'hosney mohamed',
      'islam abd elrahman', 'kamel ragab', 'mahmoud essam', 'mahmoud rabee', 'mahmoud younis',
      'mohamed elmostafa', 'mohamed shenawey', 'reda hasan abdelmaksod', 'samr nabil',
      'ahmad behiery', 'tamer lamee', 'wael zaki'
    ];

    if (['4321', '5607'].includes(String(profileRow.employee_code)) || SPECIAL_MANAGERS.includes(profileRow.employee_name?.toLowerCase())) {
      profileRow.role = 'Stakeholder';
      if (!profileRow.visible_teams || profileRow.visible_teams.length === 0) {
        profileRow.visible_teams = [...new Set((teamsRows || []).map(t => t.name))];
      }
    }

    setProfile(profileRow);
    setHierarchy(hierarchyRows || []);
    setVisibleCodes(computeVisibleEmployeeCodes(profileRow, hierarchyRows || [], teamsRows || []));
  }, []);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      if (session) {
        loadedUserId.current = session.user.id;
        loadProfileAndHierarchy(session.user.id).finally(() => setLoading(false));
      } else {
        setLoading(false);
      }
    });

    const { data: listener } = supabase.auth.onAuthStateChange((event, session) => {
      setSession(session);
      if (event === 'SIGNED_IN') {
        // supabase-js re-emits SIGNED_IN when the tab regains focus/visibility,
        // even though the user never actually signed out. If we already have
        // this user's data loaded, treat it as a no-op instead of re-showing
        // the full-screen loading spinner and re-fetching everything.
        if (loadedUserId.current === session.user.id) return;
        loadedUserId.current = session.user.id;
        setLoading(true);
        loadProfileAndHierarchy(session.user.id).finally(() => setLoading(false));
      } else if (event === 'SIGNED_OUT') {
        loadedUserId.current = null;
        setProfile(null); setHierarchy([]); setVisibleCodes([]);
        setLoading(false);
      }
    });

    return () => listener.subscription.unsubscribe();
  }, [loadProfileAndHierarchy]);

  const signInWithEmployeeCode = useCallback(async (employeeCode, password) => {
    const email = employeeCodeToEmail(employeeCode);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    return { error };
  }, []);

  const signOut = useCallback(() => supabase.auth.signOut(), []);

  const updatePassword = useCallback(async (newPassword) => {
    const { error } = await supabase.auth.updateUser({ password: newPassword });
    return { error };
  }, []);

  return (
    <AuthContext.Provider value={{
      session, profile, hierarchy, visibleCodes, loading,
      signInWithEmployeeCode, signOut, updatePassword,
    }}>
      {children}
    </AuthContext.Provider>
  );
}