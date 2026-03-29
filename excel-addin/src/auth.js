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

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

/**
 * Push current Supabase session tokens to the backend so other add-ins can pick them up.
 */
async function syncSessionToBackend(session) {
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

    const { data, error } = await supabase.auth.setSession({
      access_token: stored.access_token,
      refresh_token: stored.refresh_token,
    });
    if (error) {
      console.warn("[tsifl auth] Stored session invalid, clearing:", error.message);
      // Token expired or invalid — clear it
      await fetch(`${BACKEND_URL}/auth/clear-session`, { method: "POST" });
      return null;
    }
    // Successfully restored — sync fresh tokens back (setSession may have refreshed them)
    if (data.session) {
      await syncSessionToBackend(data.session);
    }
    return data.session?.user ?? null;
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

  // 2. No local session — try restoring from backend (shared login)
  const restored = await restoreSessionFromBackend();
  return restored ?? null;
}

/**
 * Sign in with email + password.
 * Returns { user, error }
 */
export async function signIn(email, password) {
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
 * Sign out current user — clears local and shared sessions.
 */
export async function signOut() {
  await supabase.auth.signOut();
  try {
    await fetch(`${BACKEND_URL}/auth/clear-session`, { method: "POST" });
  } catch (e) {
    console.warn("[tsifl auth] Could not clear backend session:", e);
  }
}
