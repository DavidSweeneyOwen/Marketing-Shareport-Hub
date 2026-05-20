/**
 * CheckFire Marketing Hub — Auth
 * Lightweight OAuth 2.0 PKCE — no external library required.
 */

const AUTH = {
  token: null,
  account: null,

  get authority() {
    return `https://login.microsoftonline.com/${HUB_CONFIG.tenantId}/oauth2/v2.0`;
  },

  get redirectUri() {
    return HUB_CONFIG.redirectUri || window.location.origin + window.location.pathname;
  },

  scopes: 'User.Read Sites.Read.All Files.Read.All offline_access openid profile',
};

// ── PKCE helpers ──────────────────────────────────────────────

async function generatePKCE() {
  const verifier = btoa(String.fromCharCode(...crypto.getRandomValues(new Uint8Array(32))))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  const data = new TextEncoder().encode(verifier);
  const hash = await crypto.subtle.digest('SHA-256', data);
  const challenge = btoa(String.fromCharCode(...new Uint8Array(hash)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  return { verifier, challenge };
}

// ── Sign in ───────────────────────────────────────────────────

async function signIn() {
  try {
    const { verifier, challenge } = await generatePKCE();
    const state = Math.random().toString(36).slice(2);

    sessionStorage.setItem('pkce_verifier', verifier);
    sessionStorage.setItem('pkce_state', state);

    const params = new URLSearchParams({
      client_id:             HUB_CONFIG.clientId,
      response_type:         'code',
      redirect_uri:          AUTH.redirectUri,
      scope:                 AUTH.scopes,
      code_challenge:        challenge,
      code_challenge_method: 'S256',
      state,
      prompt:                'select_account',
    });

    window.location.href = `${AUTH.authority}/authorize?${params}`;
  } catch (e) {
    showAuthError('Could not start sign-in: ' + e.message);
  }
}

// ── Handle redirect back from Microsoft ───────────────────────

async function handleRedirect() {
  const params = new URLSearchParams(window.location.search);
  const code  = params.get('code');
  const state = params.get('state');
  const error = params.get('error');

  if (error) {
    showAuthError('Sign-in error: ' + (params.get('error_description') || error));
    return false;
  }

  if (!code) return false;

  // Validate state
  if (state !== sessionStorage.getItem('pkce_state')) {
    showAuthError('Security error — please try signing in again.');
    return false;
  }

  const verifier = sessionStorage.getItem('pkce_verifier');

  try {
    const res = await fetch(`${AUTH.authority}/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id:     HUB_CONFIG.clientId,
        grant_type:    'authorization_code',
        code,
        redirect_uri:  AUTH.redirectUri,
        code_verifier: verifier,
      }),
    });

    const data = await res.json();
    if (data.error) throw new Error(data.error_description || data.error);

    AUTH.token = data.access_token;

    // Store refresh token if present
    if (data.refresh_token) sessionStorage.setItem('refresh_token', data.refresh_token);

    // Clean up URL
    window.history.replaceState({}, document.title, window.location.pathname);

    // Get user info
    await loadUserProfile();
    return true;
  } catch (e) {
    showAuthError('Token exchange failed: ' + e.message);
    return false;
  }
}

// ── Get access token (with refresh) ──────────────────────────

async function getAccessToken() {
  if (AUTH.token) return AUTH.token;

  const refresh = sessionStorage.getItem('refresh_token');
  if (!refresh) return null;

  try {
    const res = await fetch(`${AUTH.authority}/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id:     HUB_CONFIG.clientId,
        grant_type:    'refresh_token',
        refresh_token: refresh,
        scope:         AUTH.scopes,
      }),
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    AUTH.token = data.access_token;
    if (data.refresh_token) sessionStorage.setItem('refresh_token', data.refresh_token);
    return AUTH.token;
  } catch (e) {
    sessionStorage.removeItem('refresh_token');
    return null;
  }
}

// ── User profile ──────────────────────────────────────────────

async function loadUserProfile() {
  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,mail', {
      headers: { Authorization: `Bearer ${AUTH.token}` },
    });
    const user = await res.json();
    AUTH.account = user;
    showSignedIn(user);
  } catch (e) {
    showSignedIn({ displayName: 'Signed in' });
  }
}

// ── Sign out ──────────────────────────────────────────────────

function signOut() {
  AUTH.token = null;
  AUTH.account = null;
  sessionStorage.clear();
  const params = new URLSearchParams({
    client_id:   HUB_CONFIG.clientId,
    post_logout_redirect_uri: AUTH.redirectUri,
  });
  window.location.href = `${AUTH.authority}/logout?${params}`;
}

// ── Init ──────────────────────────────────────────────────────

async function initAuth() {
  if (HUB_CONFIG.clientId === 'YOUR_CLIENT_ID') {
    showDemoMode();
    return false;
  }

  // Check if returning from Microsoft login
  if (window.location.search.includes('code=')) {
    const ok = await handleRedirect();
    if (ok) return true;
  }

  // Check for existing refresh token
  if (sessionStorage.getItem('refresh_token')) {
    const token = await getAccessToken();
    if (token) {
      await loadUserProfile();
      return true;
    }
  }

  // Not signed in
  showAuthOverlay();
  return false;
}

// ── UI ────────────────────────────────────────────────────────

function showAuthOverlay() {
  const el = document.getElementById('auth-overlay');
  if (el) el.classList.remove('hidden');
}

function showSignedIn(user) {
  const overlay = document.getElementById('auth-overlay');
  if (overlay) overlay.classList.add('hidden');

  const info = document.getElementById('nav-user-info');
  if (info) info.style.display = 'flex';

  const name = user.displayName || user.mail || 'You';
  const initials = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();

  const avatar = document.getElementById('nav-user-avatar');
  const nameEl = document.getElementById('nav-user-name');
  if (avatar) avatar.textContent = initials;
  if (nameEl) nameEl.textContent = name.split(' ')[0];
}

function showDemoMode() {
  const overlay = document.getElementById('auth-overlay');
  if (overlay) overlay.classList.add('hidden');
  const banner = document.getElementById('demo-banner');
  if (banner) banner.style.display = 'flex';
  window.HUB_DEMO_MODE = true;
}

function showAuthError(msg) {
  const sub = document.querySelector('.auth-sub');
  if (sub) {
    sub.textContent = msg;
    sub.style.color = '#DC2626';
  }
  console.error('[Auth]', msg);
}

// ── Pre-build the sign-in URL so the button works as a real link ──

async function setupSignInButton() {
  try {
    const { verifier, challenge } = await generatePKCE();
    const state = Math.random().toString(36).slice(2);
    sessionStorage.setItem('pkce_verifier', verifier);
    sessionStorage.setItem('pkce_state', state);

    const params = new URLSearchParams({
      client_id:             HUB_CONFIG.clientId,
      response_type:         'code',
      redirect_uri:          AUTH.redirectUri,
      scope:                 AUTH.scopes,
      code_challenge:        challenge,
      code_challenge_method: 'S256',
      state,
      prompt:                'select_account',
    });

    const url = `https://login.microsoftonline.com/${HUB_CONFIG.tenantId}/oauth2/v2.0/authorize?${params}`;
    const btn = document.getElementById('signin-link');
    if (btn) {
      btn.href = url;
      btn.onclick = null; // remove onclick, just use the href
    }
  } catch(e) {
    console.error('Could not build sign-in URL:', e);
  }
}
