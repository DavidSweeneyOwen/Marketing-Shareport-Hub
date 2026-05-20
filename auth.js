/**
 * CheckFire Marketing Hub — Auth (lightweight PKCE, no library)
 * Uses localStorage (not sessionStorage) so state survives the Microsoft redirect
 * in corporate browser environments.
 */

const AUTH = { token: null, account: null };

function getAuthority() {
  return `https://login.microsoftonline.com/${HUB_CONFIG.tenantId}/oauth2/v2.0`;
}

function getRedirectUri() {
  return HUB_CONFIG.redirectUri || (window.location.origin + window.location.pathname);
}

// Minimal scopes — no admin consent required for sign-in
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
    const state = crypto.randomUUID
      ? crypto.randomUUID()
      : Math.random().toString(36).slice(2) + Date.now();

    // localStorage survives the browser navigation to Microsoft and back
    localStorage.setItem('hub_pkce_verifier', verifier);
    localStorage.setItem('hub_pkce_state',    state);
    localStorage.setItem('hub_pkce_time',     Date.now().toString());

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
    window.history.replaceState({}, document.title, window.location.pathname);
    showAuthOverlay();
    return false;
  }

  if (!code) return false;

  showStatus('Got code — verifying state…');

  const savedState    = localStorage.getItem('hub_pkce_state');
  const savedVerifier = localStorage.getItem('hub_pkce_verifier');
  const savedTime     = parseInt(localStorage.getItem('hub_pkce_time') || '0', 10);

  // Expire PKCE state after 10 minutes
  if (Date.now() - savedTime > 10 * 60 * 1000) {
    showStatus('Sign-in session expired — please try again.', true);
    localStorage.removeItem('hub_pkce_verifier');
    localStorage.removeItem('hub_pkce_state');
    localStorage.removeItem('hub_pkce_time');
    window.history.replaceState({}, document.title, window.location.pathname);
    showAuthOverlay();
    return false;
  }

  if (!savedState || state !== savedState) {
    showStatus(
      'State mismatch — please try again. (saved: ' + (savedState ? savedState.slice(0,8) + '…' : 'missing') +
      ' | received: ' + (state ? state.slice(0,8) + '…' : 'missing') + ')', true
    );
    window.history.replaceState({}, document.title, window.location.pathname);
    showAuthOverlay();
    return false;
  }

  showStatus('Exchanging code for token…');

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
    if (data.refresh_token) localStorage.setItem('hub_refresh', data.refresh_token);

    // Clean up PKCE state
    localStorage.removeItem('hub_pkce_verifier');
    localStorage.removeItem('hub_pkce_state');
    localStorage.removeItem('hub_pkce_time');

    window.history.replaceState({}, document.title, window.location.pathname);
    showStatus('Signed in!');
    await loadUserProfile();
    return true;

  } catch (e) {
    showStatus('Network error: ' + e.message, true);
    window.history.replaceState({}, document.title, window.location.pathname);
    showAuthOverlay();
    return false;
  }
}

// ── Token refresh ─────────────────────────────────────────────

async function getAccessToken() {
  if (AUTH.token) return AUTH.token;
  const refresh = localStorage.getItem('hub_refresh');
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
    if (data.error) throw new Error(data.error_description || data.error);
    AUTH.token = data.access_token;
    if (data.refresh_token) localStorage.setItem('hub_refresh', data.refresh_token);
    return AUTH.token;
  } catch (e) {
    localStorage.removeItem('hub_refresh');
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
  AUTH.account = null;
  localStorage.removeItem('hub_refresh');
  localStorage.removeItem('hub_pkce_verifier');
  localStorage.removeItem('hub_pkce_state');
  localStorage.removeItem('hub_pkce_time');
  window.location.href = `${getAuthority()}/logout?post_logout_redirect_uri=${encodeURIComponent(getRedirectUri())}`;
}

// ── Init ──────────────────────────────────────────────────────

async function initAuth() {
  if (HUB_CONFIG.clientId === 'YOUR_CLIENT_ID') {
    showDemoMode();
    return false;
  }

  // Handle return from Microsoft login
  if (window.location.search.includes('code=') || window.location.search.includes('error=')) {
    return await handleRedirect();
  }

  // Try silent refresh from saved token
  const refresh = localStorage.getItem('hub_refresh');
  if (refresh) {
    const token = await getAccessToken();
    if (token) {
      await loadUserProfile();
      return true;
    }
  }

  // Not signed in — show overlay
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

  // Update global nav user info if it exists
  const info = document.getElementById('nav-user-info');
  if (info) info.style.display = 'flex';

  const name     = user.displayName || user.mail || 'You';
  const initials = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();

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
  let el = document.getElementById('hub-auth-status');
  if (!el) {
    el = document.createElement('div');
    el.id = 'hub-auth-status';
    el.style.cssText = [
      'position:fixed', 'bottom:16px', 'left:50%', 'transform:translateX(-50%)',
      'padding:10px 18px', 'border-radius:10px', 'font-size:12px',
      'z-index:99999', 'max-width:90vw', 'text-align:center',
      'font-family:sans-serif', 'color:#fff', 'pointer-events:none',
    ].join(';');
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.style.background = isError ? '#DC2626' : '#111';
  el.style.display    = 'block';
  if (!isError) setTimeout(() => { if (el) el.style.display = 'none'; }, 4000);
}
