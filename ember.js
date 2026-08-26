/**
 * CheckFire Marketing Hub — Ember
 * ─────────────────────────────────────────────────────────────
 * Ember is the assistant that lives in the corner of every page.
 *
 * TWO MODES, chosen automatically from config.js:
 *
 *  1. AGENT MODE — HUB_CONFIG.ember.copilotEmbedUrl is set.
 *     The panel hosts a Copilot Studio agent published to a custom
 *     website. Copilot Studio does the AI and, importantly, honours
 *     each person's own SharePoint permissions: Ember can never show
 *     someone a document they couldn't already open. No API key ever
 *     reaches the browser. See EMBER-COPILOT-STUDIO-SETUP.md.
 *
 *  2. SEARCH MODE — the default, and what runs until the agent is
 *     built. Ember searches the Marketing Hub, Product Portal and
 *     Media Portal through Microsoft Graph with the token the hub
 *     already holds, and answers with the documents themselves,
 *     opened in-hub. No AI, no cost, no new consent — and for
 *     "where's the Kitemark certificate for fire blankets" it is
 *     genuinely the fastest answer in the building.
 *
 * Search mode is deliberately not pretending to be the agent: it says
 * what it is. When the agent lands, the panel, the launcher and the
 * keyboard shortcut all stay exactly where people learnt them.
 *
 * Load order: config.js → ui.js → auth.js → graph.js → app.js →
 *             jotform.js → ember.js
 */

const EMBER = {
  open: false,
  booted: false,
  busy: false,
  history: [],      // { role: 'you' | 'ember', html }
  driveIds: {},     // site key → drive id, resolved once
};

function _emberCfg() {
  return HUB_CONFIG.ember || {};
}

// ── Panel shell ──────────────────────────────────────────────

function toggleEmber(force) {
  const cfg = _emberCfg();
  if (cfg.enabled === false) return;

  EMBER.open = (force === undefined) ? !EMBER.open : !!force;
  const dock = document.getElementById('ember-dock');
  if (dock) dock.classList.toggle('open', EMBER.open);
  document.body.classList.toggle('ember-open', EMBER.open);

  if (EMBER.open) {
    if (!EMBER.booted) { EMBER.booted = true; _emberBoot(); }
    setTimeout(() => {
      const i = document.getElementById('ember-input');
      if (i) i.focus();
    }, 260);
  }
}

function _emberBoot() {
  const body = document.getElementById('ember-body');
  if (!body) return;
  const cfg = _emberCfg();

  // Agent mode — hand the whole panel to Copilot Studio.
  if (cfg.copilotEmbedUrl) {
    body.innerHTML = `<iframe class="ember-frame" src="${escAttr(safeUrl(cfg.copilotEmbedUrl, ''))}"
      title="${escAttr(cfg.name || 'Ember')}" frameborder="0"
      allow="microphone; clipboard-write"></iframe>`;
    const form = document.getElementById('ember-form');
    if (form) form.style.display = 'none';
    return;
  }

  _emberRender();
}

function _emberRender() {
  const body = document.getElementById('ember-body');
  if (!body) return;
  const cfg = _emberCfg();

  if (!EMBER.history.length) {
    body.innerHTML = `
      <div class="ember-welcome">
        <div class="ember-mark" aria-hidden="true">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" width="26" height="26">
            <path d="M12 2s5 4.5 5 9a5 5 0 0 1-10 0c0-1.6.6-3.1 1.4-4.3"/>
            <path d="M12 22a4 4 0 0 0 4-4c0-2-2-3.5-4-6-2 2.5-4 4-4 6a4 4 0 0 0 4 4z"/>
          </svg>
        </div>
        <h3>Ask ${escHtml(cfg.name || 'Ember')}</h3>
        <p>I'll search the Marketing Hub, the Product Portal and the Media Portal and open what I find right here.</p>
        <div class="ember-sugg">
          ${(cfg.suggestions || []).map(s =>
            `<button class="ember-chip" onclick="emberAsk('${escAttr(s).replace(/'/g, "\\'")}')">${escHtml(s)}</button>`
          ).join('')}
        </div>
        <p class="ember-mode">Search mode · answers are documents, not generated text.
          <span>The Copilot Studio agent slots into this same panel when it's ready.</span></p>
      </div>`;
    return;
  }

  body.innerHTML = EMBER.history.map(m => `
    <div class="ember-msg ${m.role}">
      ${m.role === 'ember' ? `<span class="ember-av">E</span>` : ''}
      <div class="ember-bubble">${m.html}</div>
    </div>`).join('') +
    (EMBER.busy ? `<div class="ember-msg ember"><span class="ember-av">E</span>
      <div class="ember-bubble"><span class="ember-dots"><i></i><i></i><i></i></span></div></div>` : '');

  body.scrollTop = body.scrollHeight;
}

// ── Asking ───────────────────────────────────────────────────

function emberSubmit(ev) {
  if (ev && ev.preventDefault) ev.preventDefault();
  const input = document.getElementById('ember-input');
  if (!input) return false;
  const q = input.value.trim();
  if (!q) return false;
  input.value = '';
  emberAsk(q);
  return false;
}

async function emberAsk(q) {
  if (EMBER.busy) return;
  EMBER.history.push({ role: 'you', html: escHtml(q) });
  EMBER.busy = true;
  _emberRender();

  let html;
  try {
    html = await _emberAnswer(q);
  } catch (e) {
    html = `<p>I couldn't reach SharePoint just then — ${escHtml(e.message)}. Try again in a moment.</p>`;
  }

  EMBER.busy = false;
  EMBER.history.push({ role: 'ember', html });
  _emberRender();
}

function emberClear() {
  EMBER.history = [];
  _emberRender();
}

// Search every configured site in parallel. Graph's drive search
// matches on filename and, for Office and PDF files, indexed content —
// so a question phrased in plain English still tends to land.
async function _emberAnswer(q) {
  const cfg = _emberCfg();
  const sites = cfg.searchSites || [];
  if (typeof graphFetch !== 'function' || !(window.AUTH && window.AUTH.account)) {
    return `<p>Sign in with your CheckFire account and I can search the SharePoint sites for you.</p>`;
  }

  const term = q.replace(/['"\\]/g, ' ').trim();

  const results = await Promise.all(sites.map(async s => {
    try {
      if (!EMBER.driveIds[s.key]) {
        const drive = await resolveDrive(s.url, HUB_CONFIG.documentsLibrary);
        EMBER.driveIds[s.key] = drive.id;
      }
      const id = EMBER.driveIds[s.key];
      const data = await graphFetch(
        `/drives/${id}/root/search(q='${encodeURIComponent(term)}')` +
        `?$select=id,name,size,webUrl,lastModifiedDateTime,file,folder&$top=8`);
      const rows = (data.value || [])
        .filter(f => !f.folder)
        .map(f => Object.assign({}, f, { _driveId: id }));
      return { site: s, rows };
    } catch (e) {
      return { site: s, rows: [], error: e.message };
    }
  }));

  const total = results.reduce((n, r) => n + r.rows.length, 0);

  // Hub pages worth pointing at even when no file matches.
  const shortcut = _emberShortcut(term.toLowerCase());

  if (!total) {
    return `<p>Nothing in SharePoint matched <strong>${escHtml(q)}</strong>.</p>
      <p class="ember-tip">Try the product on its own — “CO2”, “powder”, “fire blanket” — or a document type like “Kitemark” or “declaration of conformity”.</p>
      ${shortcut}`;
  }

  // Remember the files so a click can open the in-hub preview.
  EMBER.last = [];
  const groups = results.filter(r => r.rows.length).map(r => `
    <div class="ember-group">
      <div class="ember-group-head">${escHtml(r.site.label)}<span>${r.rows.length}</span></div>
      ${r.rows.map(f => {
        const i = EMBER.last.push(f) - 1;
        return `<button class="ember-hit" onclick="emberOpen(${i})">
          <span class="ember-hit-ico">${escHtml((String(f.name).split('.').pop() || 'FILE').slice(0, 4).toUpperCase())}</span>
          <span class="ember-hit-main">
            <span class="ember-hit-name">${escHtml(String(f.name).replace(/\.[a-z0-9]+$/i, ''))}</span>
            <span class="ember-hit-meta">${[humanSize(f.size), fmtSpDate(f.lastModifiedDateTime)].filter(Boolean).map(escHtml).join(' · ')}</span>
          </span>
        </button>`;
      }).join('')}
    </div>`).join('');

  return `<p>${total} document${total === 1 ? '' : 's'} for <strong>${escHtml(q)}</strong>:</p>
    ${groups}${shortcut}`;
}

// A few things worth answering with a page rather than a file.
function _emberShortcut(t) {
  const jump = (label, fn) =>
    `<button class="ember-jump" onclick="${fn};toggleEmber(false)">${escHtml(label)}</button>`;
  if (/launch|npd|new product/.test(t))          return jump('Open Product launches', "showPage('launches',1)");
  if (/campaign|activation/.test(t))             return jump('Open Campaigns', "showPage('campaigns',2)");
  if (/event|exhibition|fse|show/.test(t))       return jump('Open Trade & events', "showPage('trade',3)");
  if (/showroom|visit|book/.test(t))             return jump('Book the showroom', 'openBookingModal()');
  if (/cert|kitemark|conformity|med|mer|nta/.test(t)) return jump('Open the Product portal', 'openProductPortal()');
  return '';
}

function emberOpen(i) {
  const f = (EMBER.last || [])[i];
  if (f && typeof openDocFile === 'function') openDocFile(f);
}

// ── Keyboard ─────────────────────────────────────────────────
// Ctrl/⌘+K opens Ember from anywhere; Escape closes it. Both skipped
// while the user is typing into something else.
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === 'k') {
    e.preventDefault();
    toggleEmber(true);
    return;
  }
  if (e.key === 'Escape' && EMBER.open) toggleEmber(false);
});
