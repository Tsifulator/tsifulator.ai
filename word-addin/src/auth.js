/**
 * Tsifulator.ai — Auth Module
 * Handles login/logout via Supabase Auth.
 * User session is stored in localStorage so it persists across Excel restarts.
 */

import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

/**
 * Returns the current logged-in user, or null if not logged in.
 */
export async function getCurrentUser() {
  const { data: { session } } = await supabase.auth.getSession();
  return session?.user ?? null;
}

/**
 * Sign in with email + password.
 * Returns { user, error }
 */
export async function signIn(email, password) {
  const { data, error } = await supabase.auth.signInWithPassword({ email, password });
  return { user: data?.user ?? null, error };
}

/**
 * Sign up a new user.
 * Returns { user, error }
 */
export async function signUp(email, password) {
  const { data, error } = await supabase.auth.signUp({ email, password });
  return { user: data?.user ?? null, error };
}

/**
 * Sign out current user.
 */
export async function signOut() {
  await supabase.auth.signOut();
}
