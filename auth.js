/**
 * CheckFire Marketing Hub — Auth
 * ─────────────────────────────────────────────────────────────
 * Silent SSO first — CheckFire employees on their work devices
 * sign in automatically via their existing Microsoft session.
 * Only shows a login screen if no Microsoft session exists.
 */

const AUTH = { token: null, account: null };

function getAuthority() {
  return `https://login.microsoftonline.com/${HUB_CONFIG.tenantId}/oauth2/v2.0`;
}

function getRedirectUri() {
  return HUB_CONFIG.redirectUri || (window.location.origin + window.location.pathname);
}

// Scopes — includes SharePoint read for live list data
const SCOPES = 'User.Read openid profile offline_access Sites.Read.All Files.Read.All';

// ── PKCE ─────────────────────────────────────────────────────

async function generatePKCE() {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  const verifier = btoa(String.fromCharCode(...array))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  const data = new TextEncoder().encode(verifier);
  const hash = await crypto.subtle.digest('SHA-256', data);
  const challenge = btoa(String.fromCharCode(...new Uint8Array(hash)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  return { verifier, challenge };
}

// ── Build auth URL ────────────────────────────────────────────

async function buildAuthUrl(prompt = 'none') {
  const { verifier, challenge } = await generatePKCE();
  const state = crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2) + Date.now();

  sessionStorage.setItem('hub_pkce_verifier', verifier);
  sessionStorage.setItem('hub_pkce_state', state);
  if (prompt === 'none') sessionStorage.setItem('hub_silent_attempted', '1');

  const params = new URLSearchParams({
    client_id:             HUB_CONFIG.clientId,
    response_type:         'code',
    redirect_uri:          getRedirectUri(),
    scope:                 SCOPES,
    code_challenge:        challenge,
    code_challenge_method: 'S256',
    state,
    prompt,
  });

  return `${getAuthority()}/authorize?${params}`;
}

// ── Handle redirect back from Microsoft ──────────────────────

async function handleRedirect() {
  const params = new URLSearchParams(window.location.search);
  const code   = params.get('code');
  const state  = params.get('state');
  const error  = params.get('error');

  // Clean URL immediately
  window.history.replaceState({}, document.title, window.location.pathname);

  // Silent SSO failed — user needs to sign in interactively
  const silentErrors = ['login_required', 'interaction_required', 'consent_required', 'access_denied'];
  if (error && silentErrors.includes(error)) {
    sessionStorage.removeItem('hub_silent_attempted');
    showSignInPage();
    return false;
  }

  if (error) {
    showSignInPage('Something went wrong — please try signing in again.');
    return false;
  }

  if (!code) return false;

  const savedState    = sessionStorage.getItem('hub_pkce_state');
  const savedVerifier = sessionStorage.getItem('hub_pkce_verifier');

  if (!savedState || state !== savedState) {
    showSignInPage('Session error — please try again.');
    return false;
  }

  showStatus('Signing you in…');

  try {
    const res = await fetch(`${getAuthority()}/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id:     HUB_CONFIG.clientId,
        grant_type:    'authorization_code',
        code,
        redirect_uri:  getRedirectUri(),
        code_verifier: savedVerifier,
      }),
    });

    const data = await res.json();

    if (data.error) {
      showSignInPage('Sign-in failed — please try again.');
      return false;
    }

    AUTH.token = data.access_token;
    if (data.refresh_token) sessionStorage.setItem('hub_refresh', data.refresh_token);
    sessionStorage.removeItem('hub_pkce_verifier');
    sessionStorage.removeItem('hub_pkce_state');
    sessionStorage.removeItem('hub_silent_attempted');

    await loadUserProfile();
    return true;

  } catch (e) {
    showSignInPage('Connection error — please try again.');
    return false;
  }
}

// ── Token refresh ─────────────────────────────────────────────

async function getAccessToken() {
  if (AUTH.token) return AUTH.token;
  const refresh = sessionStorage.getItem('hub_refresh');
  if (!refresh) return null;
  try {
    const res = await fetch(`${getAuthority()}/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id:     HUB_CONFIG.clientId,
        grant_type:    'refresh_token',
        refresh_token: refresh,
        scope:         SCOPES,
      }),
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    AUTH.token = data.access_token;
    if (data.refresh_token) sessionStorage.setItem('hub_refresh', data.refresh_token);
    return AUTH.token;
  } catch (e) {
    sessionStorage.removeItem('hub_refresh');
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

function signOut() {
  AUTH.token = null;
  sessionStorage.clear();
  window.location.href = `${getAuthority()}/logout?post_logout_redirect_uri=${encodeURIComponent(getRedirectUri())}`;
}

// ── Sign-in trigger (called by button) ───────────────────────

async function signIn() {
  const url = await buildAuthUrl('login');
  window.location.href = url;
}

// ── Init ──────────────────────────────────────────────────────

async function initAuth() {
  if (HUB_CONFIG.clientId === 'YOUR_CLIENT_ID') {
    showDemoMode();
    return false;
  }

  // Coming back from Microsoft redirect
  if (window.location.search.includes('code=') || window.location.search.includes('error=')) {
    return await handleRedirect();
  }

  // Already have a valid refresh token — silently renew
  const refresh = sessionStorage.getItem('hub_refresh');
  if (refresh) {
    const token = await getAccessToken();
    if (token) { await loadUserProfile(); return true; }
  }

  // No session — try silent SSO first (uses existing Microsoft browser session)
  // Most CheckFire employees will be signed into Microsoft already
  const silentAttempted = sessionStorage.getItem('hub_silent_attempted');
  if (!silentAttempted) {
    showStatus('Signing you in…');
    const url = await buildAuthUrl('none'); // prompt=none = silent
    window.location.href = url;
    return false; // page is navigating away
  }

  // Silent failed — show sign-in page
  sessionStorage.removeItem('hub_silent_attempted');
  showSignInPage();
  return false;
}

// ── UI helpers ────────────────────────────────────────────────

function showSignInPage(message) {
  const overlay = document.getElementById('auth-overlay');
  if (!overlay) return;
  overlay.classList.remove('hidden');

  // Update message if provided
  if (message) {
    const sub = overlay.querySelector('.auth-sub');
    if (sub) sub.textContent = message;
  }
}

function showSignedIn(user) {
  document.getElementById('auth-overlay')?.classList.add('hidden');
  const info = document.getElementById('nav-user-info');
  if (info) info.style.display = 'flex';
  const name = user.displayName || user.mail || 'You';
  const initials = name.split(' ').map(w => w[0]).join('').slice(0,2).toUpperCase();
  const avatar = document.getElementById('nav-user-avatar');
  const nameEl = document.getElementById('nav-user-name');
  if (avatar) avatar.textContent = initials;
  if (nameEl)  nameEl.textContent = name.split(' ')[0];
  if (typeof loadShowroomData  === 'function') loadShowroomData();
  if (typeof loadWordPressNews === 'function') loadWordPressNews();
}

function showDemoMode() {
  document.getElementById('auth-overlay')?.classList.add('hidden');
  const banner = document.getElementById('demo-banner');
  if (banner) banner.style.display = 'flex';
  window.HUB_DEMO_MODE = true;
  if (typeof loadShowroomData  === 'function') loadShowroomData();
  if (typeof loadWordPressNews === 'function') loadWordPressNews();
}

function showStatus(msg) {
  let el = document.getElementById('hub-auth-status');
  if (!el) {
    el = document.createElement('div');
    el.id = 'hub-auth-status';
    el.style.cssText = 'position:fixed;bottom:16px;left:50%;transform:translateX(-50%);background:#0A0A0A;color:#fff;padding:10px 18px;border-radius:10px;font-size:12px;z-index:99999;max-width:90vw;text-align:center;pointer-events:none';
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.style.display = 'block';
  setTimeout(() => { if (el) el.style.display = 'none'; }, 3000);
}
