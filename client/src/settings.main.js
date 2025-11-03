import { auth } from './lib/firebaseClient';
import { onAuthStateChanged, signInWithEmailAndPassword } from 'firebase/auth';

const sharedEmail = import.meta.env.VITE_SHARED_EMAIL;
const sharedPassword = import.meta.env.VITE_SHARED_PASSWORD;

async function bootstrap() {
  if (!sharedEmail || !sharedPassword) {
    console.warn('Shared credentials missing. Loading settings in offline mode.');
    await import('./settings.js');
    return;
  }

  let sessionReady = false;
  try {
    await new Promise((resolve, reject) => {
      const unsubscribe = onAuthStateChanged(
        auth,
        (user) => {
          if (user) {
            window.currentUser = user;
            unsubscribe();
            sessionReady = true;
            resolve();
          } else {
            signInWithEmailAndPassword(auth, sharedEmail, sharedPassword)
              .then(() => {
                sessionReady = true;
                resolve();
              })
              .catch(reject);
          }
        },
        reject,
      );
    });
  } catch (error) {
    console.error('Failed to initialise Firebase session for settings', error);
  } finally {
    if (!sessionReady) {
      console.warn('Proceeding with local settings only.');
    }
    await import('./settings.js');
  }
}

bootstrap();
