import { auth } from './lib/firebaseClient';
import { onAuthStateChanged, signInWithEmailAndPassword } from 'firebase/auth';

const sharedEmail = import.meta.env.VITE_SHARED_EMAIL;
const sharedPassword = import.meta.env.VITE_SHARED_PASSWORD;

async function bootstrap() {
  if (!sharedEmail || !sharedPassword) {
    console.error('Shared credentials missing. Set VITE_SHARED_EMAIL and VITE_SHARED_PASSWORD.');
    return;
  }

  try {
    await new Promise((resolve, reject) => {
      const unsubscribe = onAuthStateChanged(
        auth,
        (user) => {
          if (user) {
            window.currentUser = user;
            unsubscribe();
            resolve();
          } else {
            signInWithEmailAndPassword(auth, sharedEmail, sharedPassword).catch(reject);
          }
        },
        reject,
      );
    });

    await import('./app.js');
  } catch (error) {
    console.error('Failed to initialise Firebase session', error);
    window.alert('Unable to connect to the shared workspace. Please check the shared credentials.');
  }
}

bootstrap();
