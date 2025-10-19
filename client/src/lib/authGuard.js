import { supabase } from './supabaseClient';

export async function requireSession({
  redirectTo = '/auth.html',
  onSession,
} = {}) {
  try {
    const { data, error } = await supabase.auth.getSession();
    if (error) throw error;

    const session = data?.session ?? null;
    if (!session) {
      window.location.replace(redirectTo);
      return null;
    }

    if (typeof onSession === 'function') {
      onSession(session);
    }

    supabase.auth.onAuthStateChange((_event, nextSession) => {
      if (!nextSession) {
        window.location.replace(redirectTo);
      }
    });

    return session;
  } catch (error) {
    console.error('Failed to resolve Supabase session', error);
    window.location.replace(redirectTo);
    return null;
  }
}
