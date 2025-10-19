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
  profileButton?.addEventListener('click', async (event) => {
    event.preventDefault();
    await supabase.auth.signOut();
  });
}

bootstrap();
