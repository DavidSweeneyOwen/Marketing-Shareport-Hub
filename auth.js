/**
 * CheckFire Marketing Hub — Auth (lightweight PKCE, no library)
 */

const AUTH = { token: null, account: null };

function getAuthority() {
  return `https://login.microsoftonline.com/${HUB_CONFIG.tenantId}/oauth2/v2.0`;
}

function getRedirectUri() {
  return HUB_CONFIG.redirectUri || (window.location.origin + window.location.pathname);
}

// Start with minimal scopes — no admin consent required
const SCOPES = 'User.Read openid profile offline_access';

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

// ── Sign in button setup ──────────────────────────────────────

async function setupSignInButton() {
  try {
    const { verifier, challenge } = await generatePKCE();
    const state = crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2) + Date.now();

    sessionStorage.setItem('hub_pkce_verifier', verifier);
    sessionStorage.setItem('hub_pkce_state', state);

    const params = new URLSearchParams({
      client_id:             HUB_CONFIG.clientId,
      response_type:         'code',
      redirect_uri:          getRedirectUri(),
      scope:                 SCOPES,
      code_challenge:        challenge,
      code_challenge_method: 'S256',
      state,
      prompt:                'select_account',
    });

    const url = `${getAuthority()}/authorize?${params}`;
    const btn = document.getElementById('signin-link');
    if (btn) { btn.href = url; btn.onclick = null; }

    showStatus('Ready to sign in');
  } catch (e) {
    showStatus('Setup error: ' + e.message, true);
  }
}

// ── Handle redirect back ──────────────────────────────────────

async function handleRedirect() {
  const params = new URLSearchParams(window.location.search);
  const code   = params.get('code');
  const state  = params.get('state');
  const error  = params.get('error');

  if (error) {
    showStatus('Microsoft error: ' + (params.get('error_description') || error), true);
    showAuthOverlay();
    return false;
  }

  if (!code) return false;

  showStatus('Got code — verifying state...');

  const savedState    = sessionStorage.getItem('hub_pkce_state');
  const savedVerifier = sessionStorage.getItem('hub_pkce_verifier');

  if (!savedState || state !== savedState) {
    showStatus('State mismatch — please try signing in again. (saved: ' + (savedState ? 'yes' : 'missing') + ')', true);
    // Clean URL and show sign-in
    window.history.replaceState({}, document.title, window.location.pathname);
    showAuthOverlay();
    return false;
  }

  showStatus('Exchanging code for token...');

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
      showStatus('Token error: ' + (data.error_description || data.error), true);
      window.history.replaceState({}, document.title, window.location.pathname);
      showAuthOverlay();
      return false;
    }

    AUTH.token = data.access_token;
    if (data.refresh_token) sessionStorage.setItem('hub_refresh', data.refresh_token);

    sessionStorage.removeItem('hub_pkce_verifier');
    sessionStorage.removeItem('hub_pkce_state');

    window.history.replaceState({}, document.title, window.location.pathname);
    showStatus('Signed in!');
    await loadUserProfile();
    return true;

  } catch (e) {
    showStatus('Fetch error: ' + e.message, true);
    window.history.replaceState({}, document.title, window.location.pathname);
    showAuthOverlay();
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

// ── Init ──────────────────────────────────────────────────────

async function initAuth() {
  if (HUB_CONFIG.clientId === 'YOUR_CLIENT_ID') {
    showDemoMode();
    return false;
  }

  if (window.location.search.includes('code=') || window.location.search.includes('error=')) {
    return await handleRedirect();
  }

  const refresh = sessionStorage.getItem('hub_refresh');
  if (refresh) {
    const token = await getAccessToken();
    if (token) { await loadUserProfile(); return true; }
  }

  showAuthOverlay();
  return false;
}

// ── UI helpers ────────────────────────────────────────────────

function showAuthOverlay() {
  const el = document.getElementById('auth-overlay');
  if (el) el.classList.remove('hidden');
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
}

function showDemoMode() {
  document.getElementById('auth-overlay')?.classList.add('hidden');
  const banner = document.getElementById('demo-banner');
  if (banner) banner.style.display = 'flex';
  window.HUB_DEMO_MODE = true;
}

function showStatus(msg, isError) {
  // Show status in the auth card so it's always visible
  let el = document.getElementById('hub-auth-status');
  if (!el) {
    el = document.createElement('div');
    el.id = 'hub-auth-status';
    el.style.cssText = 'position:fixed;bottom:16px;left:50%;transform:translateX(-50%);background:#0A0A0A;color:#fff;padding:10px 18px;border-radius:10px;font-size:12px;z-index:99999;max-width:90vw;text-align:center';
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.style.background = isError ? '#DC2626' : '#0A0A0A';
  el.style.display = 'block';
  if (!isError) setTimeout(() => { if (el) el.style.display = 'none'; }, 4000);
}
