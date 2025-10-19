import { supabase } from './lib/supabaseClient';
import { requireSession } from './lib/authGuard';

async function bootstrap() {
  const session = await requireSession({
    onSession(currentSession) {
      window.currentUser = currentSession.user;
    },
  });

  if (!session) return;

  await import('./app.js');

  const profileButton = document.getElementById('profile-btn');
  const signOutBtn = document.getElementById('signOutBtn');

  const handleSignOut = async (event) => {
    event?.preventDefault?.();
    try {
      const { error } = await supabase.auth.signOut();
      if (error) {
        console.warn('Supabase signOut warning:', error.message);
      }
    } catch (error) {
      console.error('Failed to sign out', error);
      window.alert('Unable to sign out. Redirecting to login.');
    } finally {
      window.location.href = '/auth.html?message=signed_out';
    }
  };

  profileButton?.addEventListener('click', handleSignOut);
  signOutBtn?.addEventListener('click', handleSignOut);
}

bootstrap();
