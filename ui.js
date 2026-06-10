/**
 * CheckFire Marketing Hub — UI Core
 * ─────────────────────────────────────────────────────────────
 * Navigation, toasts, tab switching, plus the shared escaping
 * helpers every renderer in the hub uses.
 *
 * Load order: config.js → ui.js → auth.js → graph.js → app.js → jotform.js
 */

// ── Escaping & URL helpers ───────────────────────────────────
// escHtml  → text content
// escAttr  → attribute values (quotes included)
// safeUrl  → href/src — only http(s) URLs survive, anything else
//            (javascript:, data:, malformed) returns the fallback
// safeCssUrl → for style="background-image:url(...)" contexts —
//            additionally encodes quotes, parens and backslashes
//            so a value can never break out of url("...")

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function escAttr(str) {
  return escHtml(str);
}

function safeUrl(url, fallback = '#') {
  if (url == null || url === '') return fallback;
  try {
    const u = new URL(String(url), window.location.href);
    if (u.protocol === 'https:' || u.protocol === 'http:') return u.href;
  } catch (_) { /* malformed */ }
  return fallback;
}

function safeCssUrl(url) {
  const safe = safeUrl(url, '');
  if (!safe) return '';
  return safe.replace(/[()'"\\]/g, m => (
    { '(': '%28', ')': '%29', "'": '%27', '"': '%22', '\\': '%5C' }[m]
  ));
}

// ── Page navigation ──────────────────────────────────────────
// Pages are <main class="page" id="page-{id}"> elements.
// CSS handles visibility/animation via .page.active.
// Signature kept as showPage(id, idx) — idx is legacy, ignored.

function showPage(id /*, idx */) {
  if (!id) return;
  const target = document.getElementById('page-' + id);
  if (!target) return;

  document.querySelectorAll('main.page').forEach(p => {
    p.classList.toggle('active', p === target);
  });

  window.scrollTo(0, 0);

  if (typeof updateNavActive === 'function') updateNavActive(id);
}

// ── Toast notifications ──────────────────────────────────────
// styles.css already defines .toast (fixed, bottom-centre) and
// .toast.show — we just create one element and reuse it.

let _hubToastTimer = null;

function showToast(msg) {
  let el = document.getElementById('hub-toast');
  if (!el) {
    el = document.createElement('div');
    el.id = 'hub-toast';
    el.className = 'toast';
    document.body.appendChild(el);
  }
  el.textContent = String(msg ?? '');

  // Force a frame so re-triggering while visible still animates
  el.classList.remove('show');
  requestAnimationFrame(() => el.classList.add('show'));

  clearTimeout(_hubToastTimer);
  _hubToastTimer = setTimeout(() => el.classList.remove('show'), 2500);
}

// ── Tab switching ────────────────────────────────────────────
// Launches page: buttons .ltab inside .ltabs, panes #tab-{name}
// Training page: same buttons, panes #ttab-{name}
// Both scoped to their own <main> so the two tab sets never clash.

function _switchTabSet(btn, paneId) {
  if (!btn) return;

  const nav = btn.closest('.ltabs');
  if (nav) {
    nav.querySelectorAll('.ltab').forEach(b => b.classList.toggle('active', b === btn));
  }

  const scope = btn.closest('main') || document;
  const pane = scope.querySelector('#' + paneId);
  if (!pane) return;

  scope.querySelectorAll('.ltab-pane').forEach(p => p.classList.toggle('active', p === pane));
}

function switchTab(btn, name) {
  _switchTabSet(btn, 'tab-' + name);
}

function switchTrainingTab(btn, name) {
  _switchTabSet(btn, 'ttab-' + name);
}

// ── Resources tab loader ─────────────────────────────────────
// Demo mode / signed-out → friendly message.
// Signed in → hand off to the Graph layer (graph.js).

function loadResourcesData() {
  const grid = document.getElementById('sp-documents-grid');
  if (!grid) return;

  const signedIn = window.AUTH && window.AUTH.account;
  if (window.HUB_DEMO_MODE || !signedIn) {
    grid.innerHTML = '<p class="prose dim">Sign in with your CheckFire account to see live documents from the SharePoint library.</p>';
    return;
  }

  if (typeof loadSharePointDocuments === 'function') {
    loadSharePointDocuments();
  }
}
