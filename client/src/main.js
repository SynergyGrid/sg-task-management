import { ensureSharedSession } from './lib/sharedSession';

async function bootstrap() {
  try {
    await ensureSharedSession();
    await import('./app.js');
  } catch (error) {
    console.error('Failed to initialise shared session', error);
    const message =
      error.message ??
      'Unable to join the shared workspace. Check VITE_SHARED_EMAIL and VITE_SHARED_PASSWORD.';
    window.alert(message);
  }
}

bootstrap();
