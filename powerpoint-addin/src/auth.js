/**
 * Tsifulator.ai — Auth Module (Shared Session)
 * Handles login/logout via Supabase Auth.
 * Sessions are synced through the backend so logging in on one add-in
 * (Excel, Word, PowerPoint) automatically logs you in everywhere.
 */

import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";
const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: false,
    storage: {
      getItem: (key) => {
        try { return localStorage.getItem(key); } catch { return null; }
      },
      setItem: (key, value) => {
        try { localStorage.setItem(key, value); } catch {}
      },
      removeItem: (key) => {
        try { localStorage.removeItem(key); } catch {}
      },
    },
  },
});

/**
 * Listen for auth state changes — sync fresh tokens to backend whenever
 * Supabase auto-refreshes the session (this is the key to staying logged in).
 */
supabase.auth.onAuthStateChange((event, session) => {
  if (session && (event === "TOKEN_REFRESHED" || event === "SIGNED_IN")) {
    syncSessionToBackend(session);
  }
});

// Proactive token refresh every 45 minutes (Improvement 2)
setInterval(async () => {
  try {
    const { data } = await supabase.auth.refreshSession();
    if (data?.session) {
      syncSessionToBackend(data.session);
    }
  } catch (e) {
    console.warn("[tsifl auth] Proactive refresh failed:", e);
  }
}, 45 * 60 * 1000);

/**
 * Push current Supabase session tokens to the backend so other add-ins can pick them up.
 */
export async function syncSessionToBackend(session) {
  if (!session) return;
  try {
    await fetch(`${BACKEND_URL}/auth/set-session`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        access_token: session.access_token,
        refresh_token: session.refresh_token,
        user_id: session.user?.id || "",
        email: session.user?.email || "",
      }),
    });
  } catch (e) {
    console.warn("[tsifl auth] Could not sync session to backend:", e);
  }
}

/**
 * Try to restore a session from the backend (another add-in already logged in).
 * Returns the user if successful, null otherwise.
 */
async function restoreSessionFromBackend() {
  try {
    const resp = await fetch(`${BACKEND_URL}/auth/get-session`);
    const { session: stored } = await resp.json();
    if (!stored || !stored.access_token || !stored.refresh_token) return null;

    // Try setSession — Supabase will use the refresh_token to get fresh tokens
    // even if the access_token is expired
    const { data, error } = await supabase.auth.setSession({
      access_token: stored.access_token,
      refresh_token: stored.refresh_token,
    });

    if (!error && data.session) {
      // Successfully restored — sync FRESH tokens back to backend
      await syncSessionToBackend(data.session);
      return data.session.user;
    }

    // setSession failed — try using JUST the refresh token to get a new session
    console.warn("[tsifl auth] setSession failed, trying refreshSession:", error?.message);
    const { data: refreshData, error: refreshError } = await supabase.auth.refreshSession({
      refresh_token: stored.refresh_token,
    });

    if (!refreshError && refreshData.session) {
      await syncSessionToBackend(refreshData.session);
      return refreshData.session.user;
    }

    // Both failed — refresh token is truly dead. DON'T clear backend session
    // because another app with a newer refresh token might still work.
    console.warn("[tsifl auth] Refresh token expired. User must re-login.");
    return null;
  } catch (e) {
    console.warn("[tsifl auth] Could not restore session from backend:", e);
    return null;
  }
}

/**
 * Returns the current logged-in user, or null if not logged in.
 * First checks local Supabase session, then tries the shared backend session.
 */
export async function getCurrentUser() {
  // 1. Check local session first (fast path)
  const { data: { session } } = await supabase.auth.getSession();
  if (session?.user) {
    // Already logged in locally — push tokens to backend so other add-ins can use them
    syncSessionToBackend(session);
    return session.user;
  }

  // 2. Try refreshing locally first (Supabase may have a valid refresh token in storage)
  const { data: refreshData } = await supabase.auth.refreshSession();
  if (refreshData?.session?.user) {
    syncSessionToBackend(refreshData.session);
    return refreshData.session.user;
  }

  // 3. No local session — try restoring from backend (shared login)
  const restored = await restoreSessionFromBackend();
  return restored ?? null;
}

/**
 * Sign in with email + password.
 * Returns { user, error }
 */
export async function signIn(email, password) {
  // Check rate limit before attempting (Improvement 6)
  try {
    const rateResp = await fetch(`${BACKEND_URL}/auth/check-login-rate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email }),
    });
    if (rateResp.status === 429) {
      return { user: null, error: { message: "Too many login attempts. Please wait 60 seconds." } };
    }
  } catch (e) { /* continue if rate check fails */ }

  const { data, error } = await supabase.auth.signInWithPassword({ email, password });
  if (data?.session) {
    // Sync to backend so other add-ins pick it up automatically
    await syncSessionToBackend(data.session);
  }
  return { user: data?.user ?? null, error };
}

/**
 * Sign up a new user.
 * Returns { user, error }
 */
export async function signUp(email, password) {
  const { data, error } = await supabase.auth.signUp({ email, password });
  if (data?.session) {
    await syncSessionToBackend(data.session);
  }
  return { user: data?.user ?? null, error };
}

/**
 * Sign out current user.
 * @param {boolean} everywhere - If true, clears backend session too (sign out everywhere).
 *                               If false, only clears local session. (Improvement 4)
 */
export async function signOut(everywhere = true) {
  await supabase.auth.signOut();
  if (everywhere) {
    try {
      await fetch(`${BACKEND_URL}/auth/clear-session`, { method: "POST" });
    } catch (e) {
      console.warn("[tsifl auth] Could not clear backend session:", e);
    }
  }
}

/**
 * Send password reset email (Improvement 5)
 */
export async function resetPassword(email) {
  const { error } = await supabase.auth.resetPasswordForEmail(email);
  return { error };
}

/**
 * Get session expiry time from JWT (Improvement 7)
 */
export function getTokenExpiry() {
  try {
    const session = supabase.auth.session?.();
    const token = session?.access_token;
    if (!token) return null;
    const payload = JSON.parse(atob(token.split('.')[1]));
    return payload.exp ? payload.exp * 1000 : null;
  } catch (e) {
    return null;
  }
}
