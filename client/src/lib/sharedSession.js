import { supabase } from './supabaseClient';

const sharedEmail = import.meta.env.VITE_SHARED_EMAIL;
const sharedPassword = import.meta.env.VITE_SHARED_PASSWORD;

export const ensureSharedSession = async () => {
  const { data, error } = await supabase.auth.getSession();
  if (error) throw error;

  if (data?.session) {
    window.currentUser = data.session.user;
    return data.session;
  }

  if (!sharedEmail || !sharedPassword) {
    throw new Error(
      'Shared credentials are not configured. Set VITE_SHARED_EMAIL and VITE_SHARED_PASSWORD.'
    );
  }

  const {
    data: signInData,
    error: signInError,
  } = await supabase.auth.signInWithPassword({
    email: sharedEmail,
    password: sharedPassword,
  });

  if (signInError) throw signInError;
  window.currentUser = signInData.session.user;
  return signInData.session;
};
