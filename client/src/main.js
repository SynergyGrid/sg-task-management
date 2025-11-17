import { auth } from './lib/firebaseClient';
import { onAuthStateChanged, signInWithEmailAndPassword } from 'firebase/auth';

const sharedEmail = import.meta.env.VITE_SHARED_EMAIL ?? '';
const sharedPassword = import.meta.env.VITE_SHARED_PASSWORD ?? '';

const loginForm = document.getElementById('loginForm');
const loginScreen = document.getElementById('loginScreen');
const loginEmail = document.getElementById('loginEmail');
const loginPassword = document.getElementById('loginPassword');
const loginSubmit = document.getElementById('loginSubmit');
const loginStatus = document.getElementById('loginStatus');

let loginOverlayVisible = false;

const showLoginOverlay = () => {
  if (!loginScreen || loginOverlayVisible) return;
  loginScreen.classList.remove('hidden');
  loginScreen.setAttribute('aria-hidden', 'false');
  loginOverlayVisible = true;
  loginEmail?.focus();
};

const hideLoginOverlay = () => {
  if (!loginScreen) return;
  loginScreen.classList.add('hidden');
  loginScreen.setAttribute('aria-hidden', 'true');
  loginOverlayVisible = false;
};

const authReadyPromise = new Promise((resolve, reject) => {
  const unsubscribe = onAuthStateChanged(
    auth,
    (user) => {
      if (user) {
        window.currentUser = user;
        unsubscribe();
        resolve(user);
      } else {
        showLoginOverlay();
        setStatus('');
      }
    },
    (error) => {
      unsubscribe();
      reject(error);
    },
  );
});

let appLaunchPromise = null;

const setStatus = (message = '', variant = 'neutral') => {
  if (!loginStatus) return;
  loginStatus.textContent = message;
  loginStatus.classList.remove('text-red-500', 'text-slate-500');
  loginStatus.setAttribute('aria-hidden', message ? 'false' : 'true');
  if (!message) {
    loginStatus.classList.add('text-slate-500');
    return;
  }
  loginStatus.classList.add(variant === 'error' ? 'text-red-500' : 'text-slate-500');
};

const launchWorkspace = async () => {
  if (appLaunchPromise) return appLaunchPromise;
  appLaunchPromise = (async () => {
    try {
      await authReadyPromise;
      if (loginScreen) {
        loginScreen.classList.add('hidden');
        loginScreen.remove();
      }
      await import('./app.js');
    } finally {
      // keep appLaunchPromise in place so we don't re-import
    }
  })();
  return appLaunchPromise;
};

const handleLoginError = (error) => {
  console.error('Sign-in error', error);
  const message =
    error?.message?.replace('Firebase: ', '').replace('auth/', '') || 'Unable to sign in with those credentials';
  setStatus(message, 'error');
};

const attemptSignIn = async (event) => {
  event?.preventDefault();
  if (!loginForm || !loginSubmit) return;
  const email = (loginEmail?.value ?? '').trim();
  const password = (loginPassword?.value ?? '').trim();
  if (!email || !password) {
    setStatus('Please enter the workspace email and password.', 'error');
    return;
  }

  showLoginOverlay();
  loginSubmit.disabled = true;
  setStatus('Connecting...');
  try {
    await signInWithEmailAndPassword(auth, email, password);
    await launchWorkspace();
  } catch (error) {
    handleLoginError(error);
  } finally {
    loginSubmit.disabled = false;
  }
};

if (loginEmail && sharedEmail) {
  loginEmail.value = sharedEmail;
}

if (loginPassword && sharedPassword) {
  loginPassword.placeholder = 'Shared workspace password';
}

loginForm?.addEventListener('submit', attemptSignIn);
loginSubmit?.addEventListener('click', attemptSignIn);

authReadyPromise
  .then(() => {
    launchWorkspace().catch((error) => {
      console.error('Unable to initialize workspace', error);
      setStatus('Unable to open workspace right now. Please refresh.', 'error');
    });
  })
  .catch((error) => {
    console.error('Firebase initialization failed', error);
    setStatus('Unable to connect to Firebase right now.', 'error');
  });
