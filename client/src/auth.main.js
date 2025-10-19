import { supabase } from './lib/supabaseClient';

const authForm = document.getElementById('authForm');
const emailInput = document.getElementById('email');
const passwordInput = document.getElementById('password');
const submitButton = document.getElementById('submitButton');
const authMessage = document.getElementById('authMessage');
const toggleModeButton = document.getElementById('toggleMode');

let mode = 'sign-in';

const updateUiForMode = () => {
  if (!toggleModeButton || !submitButton) return;
  if (mode === 'sign-in') {
    submitButton.textContent = 'Sign in';
    toggleModeButton.textContent = 'Need an account? Create one';
  } else {
    submitButton.textContent = 'Create workspace account';
    toggleModeButton.textContent = 'Already have an account? Sign in';
  }
  authMessage.textContent = '';
  authMessage.classList.remove('text-red-300', 'text-emerald-200');
};

toggleModeButton?.addEventListener('click', () => {
  mode = mode === 'sign-in' ? 'sign-up' : 'sign-in';
  updateUiForMode();
});

const setLoading = (isLoading) => {
  submitButton.disabled = isLoading;
  submitButton.classList.toggle('opacity-70', isLoading);
  submitButton.textContent = isLoading
    ? mode === 'sign-in'
      ? 'Signing in...'
      : 'Creating account...'
    : mode === 'sign-in'
    ? 'Sign in'
    : 'Create workspace account';
};

const redirectToApp = () => {
  window.location.href = '/';
};

async function handleSubmit(event) {
  event.preventDefault();
  const email = emailInput.value.trim();
  const password = passwordInput.value.trim();

  if (!email || !password) {
    authMessage.textContent = 'Email and password are required.';
    authMessage.classList.add('text-red-300');
    return;
  }

  setLoading(true);
  authMessage.textContent = '';
  authMessage.classList.remove('text-red-300', 'text-emerald-200');

  try {
    if (mode === 'sign-in') {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) throw error;
      authMessage.textContent = 'Success! Redirecting...';
      authMessage.classList.add('text-emerald-200');
      redirectToApp();
    } else {
      const { error } = await supabase.auth.signUp({
        email,
        password,
        options: {
          data: {},
        },
      });
      if (error) throw error;
      authMessage.textContent =
        'Account created! Check your email to confirm, then sign in.';
      authMessage.classList.add('text-emerald-200');
      mode = 'sign-in';
      updateUiForMode();
    }
  } catch (error) {
    authMessage.textContent = error.message ?? 'Something went wrong. Try again.';
    authMessage.classList.add('text-red-300');
  } finally {
    setLoading(false);
  }
}

authForm?.addEventListener('submit', handleSubmit);

supabase.auth.getSession().then(({ data }) => {
  if (data?.session) {
    redirectToApp();
  }
});

supabase.auth.onAuthStateChange((_event, session) => {
  if (session) {
    redirectToApp();
  }
});

updateUiForMode();
