/**
 * CheckFire Marketing Hub — Auth (MSAL.js)
 * ─────────────────────────────────────────────────────────────
 * Silent SSO first: ssoSilent without a hint only works reliably
 * when the browser holds exactly one Microsoft session, so we
 * store the last signed-in username as a loginHint — making
 * silent sign-in deterministic on work devices from the second
 * visit onwards. First-ever visit falls back to one click.
 *
 * Public surface (unchanged from the previous implementation):
 *   initAuth() · getAccessToken() · signIn() · signOut()
 *   window.AUTH = { token, account }
 *
 * Requires msal-browser.min.js loaded BEFORE this file.
 */

const AUTH = { token: null, account: null };
window.AUTH = AUTH; // const declarations don't attach to window — do it explicitly

const SCOPES = ['User.Read', 'Sites.Read.All', 'Files.Read.All'];
const LOGIN_HINT_KEY = 'hub_login_hint';

let _msalApp = null;

function getMsal() {
  if (_msalApp) return _msalApp;
  _msalApp = new msal.PublicClientApplication({
    auth: {
      clientId:    HUB_CONFIG.clientId,
      authority:   'https://login.microsoftonline.com/' + HUB_CONFIG.tenantId,
      redirectUri: HUB_CONFIG.redirectUri || (window.location.origin + window.location.pathname),
    },
    cache: { cacheLocation: 'sessionStorage' },
  });
  return _msalApp;
}

// ── Init ──────────────────────────────────────────────────────

async function initAuth() {
  // Demo mode — config not filled in yet
  if (!HUB_CONFIG.clientId || HUB_CONFIG.clientId === 'YOUR_CLIENT_ID') {
    showDemoMode();
    return false;
  }

  if (typeof msal === 'undefined') {
    console.error('[Auth] MSAL not loaded — check the msal-browser.min.js script tag in index.html');
    showSignInPage('Sign-in is unavailable — please contact IT.');
    return false;
  }

  const app = getMsal();
  if (typeof app.initialize === 'function') await app.initialize();

  // 1. Returning from a redirect sign-in?
  try {
    const result = await app.handleRedirectPromise();
    if (result && result.account) {
      return finishSignIn(result.account, result.accessToken);
    }
  } catch (e) {
    console.warn('[Auth] Redirect handling failed:', e.message);
    showSignInPage('Sign-in failed — please try again.');
    return false;
  }

  // 2. Account already cached this session?
  const accounts = app.getAllAccounts();
  if (accounts.length > 0) {
    return finishSignIn(accounts[0]);
  }

  // 3. Silent SSO — no clicks if the device already has a Microsoft session
  let hint = null;
  try { hint = localStorage.getItem(LOGIN_HINT_KEY); } catch (_) {}
  try {
    const result = await app.ssoSilent({
      scopes: SCOPES,
      ...(hint ? { loginHint: hint } : {}),
    });
    if (result && result.account) {
      return finishSignIn(result.account, result.accessToken);
    }
  } catch (_) {
    // InteractionRequired — expected on first visit or multi-session devices
  }

  // 4. Nothing worked silently — show the sign-in page
  showSignInPage();
  return false;
}

async function finishSignIn(account, accessToken) {
  const app = getMsal();
  app.setActiveAccount(account);

  AUTH.account = {
    displayName: account.name || account.username || 'Signed in',
    mail:        account.username || '',
  };

  try { localStorage.setItem(LOGIN_HINT_KEY, account.username || ''); } catch (_) {}

  AUTH.token = accessToken || await getAccessToken();

  showSignedIn(AUTH.account);
  return true;
}

// ── Token acquisition ─────────────────────────────────────────

async function getAccessToken() {
  if (window.HUB_DEMO_MODE) return null;
  if (typeof msal === 'undefined') return null;

  const app = getMsal();
  const account = app.getActiveAccount() || app.getAllAccounts()[0];
  if (!account) return null;

  try {
    const result = await app.acquireTokenSilent({ scopes: SCOPES, account });
    AUTH.token = result.accessToken;
    return result.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      // Session genuinely expired — one redirect re-establishes it
      await app.acquireTokenRedirect({ scopes: SCOPES });
    } else {
      console.warn('[Auth] Token acquisition failed:', e.message);
    }
    return null;
  }
}

// ── Sign in / out ─────────────────────────────────────────────

function signIn() {
  if (typeof msal === 'undefined') return;
  getMsal().loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
}

function signOut() {
  AUTH.token = null;
  AUTH.account = null;
  try { localStorage.removeItem(LOGIN_HINT_KEY); } catch (_) {}
  if (typeof msal === 'undefined') { window.location.reload(); return; }
  getMsal().logoutRedirect({
    postLogoutRedirectUri: HUB_CONFIG.redirectUri || (window.location.origin + window.location.pathname),
  });
}

// ── UI helpers ────────────────────────────────────────────────

function showSignInPage(message) {
  const overlay = document.getElementById('auth-overlay');
  if (!overlay) return;
  overlay.classList.remove('hidden');

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
  const initials = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
  const avatar = document.getElementById('nav-user-avatar');
  const nameEl = document.getElementById('nav-user-name');
  if (avatar) avatar.textContent = initials;
  if (nameEl) nameEl.textContent = name.split(' ')[0];

  // Kick off live data
  if (typeof loadShowroomData    === 'function') loadShowroomData();
  if (typeof loadWordPressNews   === 'function') loadWordPressNews();
  if (typeof loadSharePointData  === 'function') loadSharePointData();
}

function showDemoMode() {
  document.getElementById('auth-overlay')?.classList.add('hidden');
  const banner = document.getElementById('demo-banner');
  if (banner) banner.style.display = 'flex';
  window.HUB_DEMO_MODE = true;
  if (typeof loadShowroomData  === 'function') loadShowroomData();
  if (typeof loadWordPressNews === 'function') loadWordPressNews();
}
