/**
 * CheckFire Marketing Hub — Auth (MSAL.js)
 * ─────────────────────────────────────────
 * Handles Microsoft sign-in and provides access tokens for Graph API calls.
 * Uses MSAL.js 3.x (loaded via CDN in index.html).
 */

let msalInstance = null;
let currentAccount = null;

const GRAPH_SCOPES = [
  'User.Read',
  'Sites.Read.All',
  'Files.Read.All',
];

function getMsalConfig() {
  const redirectUri = HUB_CONFIG.redirectUri || window.location.origin + '/';
  return {
    auth: {
      clientId: HUB_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${HUB_CONFIG.tenantId}`,
      redirectUri,
    },
    cache: {
      cacheLocation: 'sessionStorage',
      storeAuthStateInCookie: false,
    },
  };
}

async function initAuth() {
  if (HUB_CONFIG.clientId === 'YOUR_CLIENT_ID') {
    // Config not yet filled in — show demo mode banner
    showDemoMode();
    return false;
  }

  msalInstance = new msal.PublicClientApplication(getMsalConfig());
  await msalInstance.initialize();

  // Handle redirect response (user coming back after login)
  await msalInstance.handleRedirectPromise();

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    currentAccount = accounts[0];
    showSignedIn(currentAccount);
    return true;
  }

  // Not signed in — show auth overlay
  showAuthOverlay();
  return false;
}

async function signIn() {
  if (!msalInstance) return;
  try {
    await msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
  } catch (e) {
    console.error('Sign-in failed:', e);
    showToast('Sign-in failed — please try again');
  }
}

async function getAccessToken() {
  if (!msalInstance || !currentAccount) return null;
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: currentAccount,
    });
    return result.accessToken;
  } catch (e) {
    // Silent failed — try interactive
    try {
      const result = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
      return result.accessToken;
    } catch (e2) {
      console.error('Token acquisition failed:', e2);
      return null;
    }
  }
}

function signOut() {
  if (!msalInstance || !currentAccount) return;
  msalInstance.logoutRedirect({ account: currentAccount });
}

// ── UI helpers ────────────────────────────────────────────────

function showAuthOverlay() {
  document.getElementById('auth-overlay').classList.remove('hidden');
}

function showSignedIn(account) {
  document.getElementById('auth-overlay').classList.add('hidden');
  const name = account.name || account.username || 'You';
  const initials = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
  const avatar = document.getElementById('nav-user-avatar');
  const nameEl = document.getElementById('nav-user-name');
  if (avatar) avatar.textContent = initials;
  if (nameEl) nameEl.textContent = name.split(' ')[0];
}

function showDemoMode() {
  document.getElementById('auth-overlay').classList.add('hidden');
  const banner = document.getElementById('demo-banner');
  if (banner) banner.style.display = 'flex';
  // In demo mode, data loads from static fallback content
  window.HUB_DEMO_MODE = true;
}
