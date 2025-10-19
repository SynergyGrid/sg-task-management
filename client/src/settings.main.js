import { ensureSharedSession } from './lib/sharedSession';

async function bootstrap() {
  try {
    await ensureSharedSession();
    await import('./settings.js');
  } catch (error) {
    console.error('Failed to initialise shared session for settings', error);
  }
}

bootstrap();
