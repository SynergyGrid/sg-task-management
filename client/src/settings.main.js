import { requireSession } from './lib/authGuard';

async function bootstrap() {
  const session = await requireSession({
    onSession(currentSession) {
      window.currentUser = currentSession.user;
    },
  });

  if (!session) return;

  await import('./settings.js');
}

bootstrap();
