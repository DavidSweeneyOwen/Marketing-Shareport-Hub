/**
 * CheckFire Marketing Hub — Ember
 * ─────────────────────────────────────────────────────────────
 * Ember is the assistant in the corner of every page.
 *
 * THREE MODES, picked automatically from config.js:
 *
 *  1. CLAUDE — HUB_CONFIG.ember.aiProxyUrl is set. A real
 *     conversation. See the split below; this is the interesting one.
 *
 *  2. COPILOT STUDIO — copilotEmbedUrl is set and aiProxyUrl isn't.
 *     Microsoft's agent, embedded whole. EMBER-COPILOT-STUDIO-SETUP.md.
 *
 *  3. SEARCH — neither set. No AI: Ember searches and hands back the
 *     documents themselves. Costs nothing and still beats hunting
 *     through SharePoint, so it's a sane place to sit.
 *
 * ── How the Claude mode keeps CheckFire's documents private ──
 *
 * RETRIEVAL happens HERE, in the browser, using the signed-in user's
 * own Microsoft token:
 *   · Ember searches the three SharePoint sites through Graph
 *   · it picks the most promising documents
 *   · for each one it asks Graph for a downloadUrl — a short-lived,
 *     pre-authenticated link to that single file
 *
 * GENERATION happens in the checkfire-ai Function app, which holds the
 * Anthropic key (a static site can never hold an API key — see the
 * June 2026 Jotform leak). The browser sends the question, the
 * conversation so far, and those short-lived links. The Function reads
 * the files, extracts their text, asks Claude, and returns the answer.
 *
 * What that buys:
 *   · no CheckFire document is ever indexed or stored on any server
 *   · the user's Microsoft token never leaves the browser
 *   · Ember can only ever see documents the person asking could
 *     already open — because only they could produce those links
 *
 * Load order: config.js → ui.js → auth.js → graph.js → app.js →
 *             jotform.js → ember.js
 */

const EMBER = {
  open: false,
  booted: false,
  busy: false,
  turns: [],        // { role:'user'|'assistant', text }
  view: [],         // what's on screen: { role, html }
  driveIds: {},
  lastDocs: [],
};

function _emberCfg() { return HUB_CONFIG.ember || {}; }
function _emberAiUrl() { return String(_emberCfg().aiProxyUrl || '').replace(/\/+$/, ''); }
function _emberMode() {
  if (_emberAiUrl()) return 'claude';
  if (_emberCfg().copilotEmbedUrl) return 'copilot';
  return 'search';
}

// ── Panel shell ──────────────────────────────────────────────

function toggleEmber(force) {
  if (_emberCfg().enabled === false) return;

  EMBER.open = (force === undefined) ? !EMBER.open : !!force;
  const dock = document.getElementById('ember-dock');
  if (dock) dock.classList.toggle('open', EMBER.open);
  document.body.classList.toggle('ember-open', EMBER.open);

  if (EMBER.open) {
    if (!EMBER.booted) { EMBER.booted = true; _emberBoot(); }
    setTimeout(() => { const i = document.getElementById('ember-input'); if (i) i.focus(); }, 260);
  }
}

function _emberBoot() {
  const body = document.getElementById('ember-body');
  if (!body) return;

  if (_emberMode() === 'copilot') {
    body.innerHTML = `<iframe class="ember-frame" src="${escAttr(safeUrl(_emberCfg().copilotEmbedUrl, ''))}"
      title="${escAttr(_emberCfg().name || 'Ember')}" frameborder="0"
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

  if (!EMBER.view.length) {
    const mode = _emberMode();
    body.innerHTML = `
      <div class="ember-welcome">
        <div class="ember-mark" aria-hidden="true">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" width="26" height="26">
            <path d="M12 2s5 4.5 5 9a5 5 0 0 1-10 0c0-1.6.6-3.1 1.4-4.3"/>
            <path d="M12 22a4 4 0 0 0 4-4c0-2-2-3.5-4-6-2 2.5-4 4-4 6a4 4 0 0 0 4 4z"/>
          </svg>
        </div>
        <h3>Ask ${escHtml(cfg.name || 'Ember')}</h3>
        <p>${mode === 'claude'
            ? 'Ask me anything about our products, paperwork, campaigns or training and I’ll read what we’ve got and answer properly.'
            : 'I’ll search the Marketing Hub, the Product Portal and the Media Portal and open what I find right here.'}</p>
        <div class="ember-sugg">
          ${(cfg.suggestions || []).map(s =>
            `<button class="ember-chip" onclick="emberAsk(${JSON.stringify(s).replace(/"/g, '&quot;')})">${escHtml(s)}</button>`
          ).join('')}
        </div>
        ${mode === 'claude' ? '' : `<p class="ember-mode">Search mode — answers are documents, not conversation.
          <span>Set the AI proxy in config.js to switch this on properly.</span></p>`}
      </div>`;
    return;
  }

  body.innerHTML = EMBER.view.map(m => `
    <div class="ember-msg ${m.role === 'user' ? 'you' : 'ember'}">
      ${m.role === 'user' ? '' : `<span class="ember-av">E</span>`}
      <div class="ember-bubble">${m.html}</div>
    </div>`).join('') +
    (EMBER.busy ? `<div class="ember-msg ember"><span class="ember-av">E</span>
      <div class="ember-bubble"><span class="ember-dots"><i></i><i></i><i></i></span>
      <span class="ember-status" id="ember-status"></span></div></div>` : '');

  body.scrollTop = body.scrollHeight;
}

function _emberStatus(msg) {
  const el = document.getElementById('ember-status');
  if (el) el.textContent = msg;
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
  EMBER.view.push({ role: 'user', html: escHtml(q) });
  EMBER.turns.push({ role: 'user', text: q });
  EMBER.busy = true;
  _emberRender();

  let html;
  try {
    html = _emberMode() === 'claude' ? await _emberClaude(q) : await _emberSearchOnly(q);
  } catch (e) {
    html = `<p>Something went wrong there — ${escHtml(e.message)}. Try again in a moment.</p>`;
  }

  EMBER.busy = false;
  EMBER.view.push({ role: 'assistant', html });
  _emberRender();
}

function emberClear() {
  EMBER.view = [];
  EMBER.turns = [];
  EMBER.lastDocs = [];
  _emberRender();
}

// ── Retrieval (browser side, user's own permissions) ─────────

const EMBER_READABLE = /\.(pdf|docx?|pptx?|txt|md|csv|rtf)$/i;

async function _emberSearch(term, perSite) {
  const sites = _emberCfg().searchSites || [];
  const clean = String(term || '').replace(/['"\\]/g, ' ').trim();
  if (!clean) return [];

  const results = await Promise.all(sites.map(async s => {
    try {
      if (!EMBER.driveIds[s.key]) {
        const drive = await resolveDrive(s.url, HUB_CONFIG.documentsLibrary);
        EMBER.driveIds[s.key] = drive.id;
      }
      const id = EMBER.driveIds[s.key];
      const data = await graphFetch(
        `/drives/${id}/root/search(q='${encodeURIComponent(clean)}')` +
        `?$select=id,name,size,webUrl,lastModifiedDateTime,file,folder&$top=${perSite || 8}`);
      return (data.value || [])
        .filter(f => !f.folder)
        .map(f => Object.assign({}, f, { _driveId: id, _site: s.label }));
    } catch (_) {
      return [];
    }
  }));
  return results.flat();
}

// Search on the question, and again on its most distinctive words —
// Graph's drive search is fussy about long natural-language strings.
async function _emberFind(question) {
  const stop = new Set(('a an the is are was were do does did what which who whom whose where when why how ' +
    'for of on in to from with and or our we i you it that this can could should would have has had ' +
    'please tell me about need any there their they').split(' '));
  const words = question.toLowerCase().replace(/[^a-z0-9 ]/g, ' ').split(/\s+/)
    .filter(w => w.length > 2 && !stop.has(w));

  const queries = [question];
  if (words.length) queries.push(words.slice(0, 5).join(' '));
  if (words.length > 1) queries.push(words[0]);

  const seen = new Set();
  const out = [];
  for (const q of queries) {
    const rows = await _emberSearch(q, 8);
    for (const r of rows) {
      const k = r._driveId + '|' + r.id;
      if (seen.has(k)) continue;
      seen.add(k);
      out.push(r);
    }
    if (out.length >= 20) break;
  }
  return out;
}

// What the hub already knows, handed over as plain text so Ember can
// answer "what's launching next" or "when's the next training" without
// reading a single document.
async function _emberHubContext() {
  const bits = [];
  const say = (label, rows, fn) => {
    if (!rows || !rows.length) return;
    bits.push(label + ':\n' + rows.slice(0, 12).map(fn).join('\n'));
  };
  try {
    const [launches, campaigns] = await Promise.all([
      fetchListItems(HUB_CONFIG.lists.launches).catch(() => []),
      fetchListItems(HUB_CONFIG.lists.campaigns).catch(() => []),
    ]);
    say('Product launches (from the hub)', launches,
      f => `- ${f.Title || 'Untitled'} — status ${f.Status || 'not set'}${f.LaunchDate ? ', launch date ' + fmtSpDate(f.LaunchDate) : ''}${productCodes(f).length ? ', codes ' + productCodes(f).join(', ') : ''}`);
    say('Campaigns (from the hub)', campaigns,
      f => `- ${f.Title || 'Untitled'} — status ${f.Status || 'not set'}${f.StartDate ? ', from ' + fmtSpDate(f.StartDate) : ''}${f.EndDate ? ' to ' + fmtSpDate(f.EndDate) : ''}`);
    if (typeof TRAIN === 'object' && TRAIN.sessions && TRAIN.sessions.length) {
      say('Upcoming training sessions (from the hub)', TRAIN.sessions,
        s => `- ${s.Title} on ${fmtSpDate(s.TrainingDate)}${s.Trainer ? ' with ' + s.Trainer : ''}${s.Location ? ' at ' + s.Location : ''}`);
    }
  } catch (_) { /* context is a bonus, never a blocker */ }
  return bits.join('\n\n');
}

// ── Claude mode ──────────────────────────────────────────────

async function _emberClaude(question) {
  if (!(window.AUTH && window.AUTH.account)) {
    return `<p>Sign in with your CheckFire account first and I can read our documents for you.</p>`;
  }

  _emberStatus('looking through SharePoint…');
  const found = await _emberFind(question);

  const cfg = _emberCfg();
  const maxDocs = cfg.maxDocs || 8;

  // Only documents worth reading — no images or videos.
  const readable = found.filter(f => EMBER_READABLE.test(f.name || '')).slice(0, maxDocs);

  _emberStatus(readable.length ? `reading ${readable.length} document${readable.length === 1 ? '' : 's'}…` : 'thinking…');

  // A downloadUrl is short-lived and pre-authenticated: it is a link to
  // ONE file the signed-in user can already open, and it is the only
  // thing about that file that leaves the browser.
  const docs = [];
  await Promise.all(readable.map(async f => {
    try {
      const meta = await graphFetch(
        `/drives/${f._driveId}/items/${f.id}?$select=id,name,size,@microsoft.graph.downloadUrl`);
      const url = meta && meta['@microsoft.graph.downloadUrl'];
      if (!url) return;
      docs.push({ name: f.name, site: f._site, url, size: f.size });
    } catch (_) { /* skip it */ }
  }));

  const hubContext = await _emberHubContext();
  EMBER.lastDocs = found.slice(0, 12);

  _emberStatus('thinking…');

  const res = await fetch(_emberAiUrl() + '/chat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      messages: EMBER.turns.slice(-12),
      docs,
      context: hubContext,
      fileList: found.slice(0, 20).map(f => ({ name: f.name, site: f._site })),
      maxCharsPerDoc: cfg.maxCharsPerDoc || 6000,
    }),
  });

  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    if (res.status === 404) throw new Error('the AI proxy isn’t deployed yet');
    throw new Error(`the AI proxy returned ${res.status}${detail ? ' — ' + detail.slice(0, 160) : ''}`);
  }

  const data = await res.json();
  const answer = String(data.text || data.answer || '').trim();
  if (!answer) throw new Error('the AI proxy sent an empty answer');

  EMBER.turns.push({ role: 'assistant', text: answer });

  // Only show source cards for documents Ember actually used, when the
  // proxy tells us; otherwise the best few it was given.
  const usedNames = Array.isArray(data.used) ? data.used.map(String) : null;
  const sources = usedNames
    ? EMBER.lastDocs.filter(f => usedNames.some(n => String(f.name).toLowerCase() === n.toLowerCase()))
    : EMBER.lastDocs.slice(0, 4);

  return _emberMarkdown(answer) + _emberSources(sources) + _emberShortcut(question.toLowerCase());
}

// A small, deliberately boring markdown renderer — bold, bullets,
// paragraphs and inline code. Everything is escaped first, so nothing
// the model returns can inject markup.
function _emberMarkdown(text) {
  const esc = escHtml(text);
  const lines = esc.split('\n');
  let out = '', list = null;

  const inline = s => s
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/`([^`]+)`/g, '<code>$1</code>');

  for (const raw of lines) {
    const line = raw.trim();
    const bullet = line.match(/^[-*•]\s+(.*)$/);
    const number = line.match(/^(\d+)[.)]\s+(.*)$/);

    if (bullet) {
      if (list !== 'ul') { if (list) out += `</${list}>`; out += '<ul class="ember-ul">'; list = 'ul'; }
      out += `<li>${inline(bullet[1])}</li>`;
    } else if (number) {
      if (list !== 'ol') { if (list) out += `</${list}>`; out += '<ol class="ember-ul">'; list = 'ol'; }
      out += `<li>${inline(number[2])}</li>`;
    } else {
      if (list) { out += `</${list}>`; list = null; }
      if (line) out += `<p>${inline(line)}</p>`;
    }
  }
  if (list) out += `</${list}>`;
  return out || `<p>${esc}</p>`;
}

function _emberSources(files) {
  if (!files || !files.length) return '';
  return `<div class="ember-group">
    <div class="ember-group-head">Sources<span>${files.length}</span></div>
    ${files.map(f => {
      const i = EMBER.lastDocs.indexOf(f);
      return `<button class="ember-hit" onclick="emberOpen(${i})">
        <span class="ember-hit-ico">${escHtml((String(f.name).split('.').pop() || 'FILE').slice(0, 4).toUpperCase())}</span>
        <span class="ember-hit-main">
          <span class="ember-hit-name">${escHtml(String(f.name).replace(/\.[a-z0-9]+$/i, ''))}</span>
          <span class="ember-hit-meta">${escHtml([f._site, humanSize(f.size), fmtSpDate(f.lastModifiedDateTime)].filter(Boolean).join(' · '))}</span>
        </span>
      </button>`;
    }).join('')}
  </div>`;
}

// ── Search mode (no AI configured) ───────────────────────────

async function _emberSearchOnly(question) {
  if (!(window.AUTH && window.AUTH.account)) {
    return `<p>Sign in with your CheckFire account and I can search the SharePoint sites for you.</p>`;
  }
  const found = await _emberFind(question);
  EMBER.lastDocs = found.slice(0, 20);
  const shortcut = _emberShortcut(question.toLowerCase());

  if (!found.length) {
    return `<p>Nothing in SharePoint matched <strong>${escHtml(question)}</strong>.</p>
      <p class="ember-tip">Try the product on its own — “CO2”, “powder”, “fire blanket” — or a document type like “Kitemark” or “declaration of conformity”.</p>
      ${shortcut}`;
  }

  const bySite = {};
  EMBER.lastDocs.forEach(f => { (bySite[f._site] = bySite[f._site] || []).push(f); });

  const groups = Object.keys(bySite).map(site => `
    <div class="ember-group">
      <div class="ember-group-head">${escHtml(site)}<span>${bySite[site].length}</span></div>
      ${bySite[site].map(f => {
        const i = EMBER.lastDocs.indexOf(f);
        return `<button class="ember-hit" onclick="emberOpen(${i})">
          <span class="ember-hit-ico">${escHtml((String(f.name).split('.').pop() || 'FILE').slice(0, 4).toUpperCase())}</span>
          <span class="ember-hit-main">
            <span class="ember-hit-name">${escHtml(String(f.name).replace(/\.[a-z0-9]+$/i, ''))}</span>
            <span class="ember-hit-meta">${escHtml([humanSize(f.size), fmtSpDate(f.lastModifiedDateTime)].filter(Boolean).join(' · '))}</span>
          </span>
        </button>`;
      }).join('')}
    </div>`).join('');

  return `<p>${EMBER.lastDocs.length} document${EMBER.lastDocs.length === 1 ? '' : 's'} for <strong>${escHtml(question)}</strong>:</p>
    ${groups}${shortcut}`;
}

// A few things worth answering with a page rather than a file.
function _emberShortcut(t) {
  const jump = (label, fn) =>
    `<button class="ember-jump" onclick="${fn};toggleEmber(false)">${escHtml(label)}</button>`;
  if (/launch|npd|new product/.test(t))                return jump('Open Product launches', "showPage('launches',1)");
  if (/campaign|activation/.test(t))                   return jump('Open Campaigns', "showPage('campaigns',2)");
  if (/training|course|session/.test(t))               return jump('Open Trade, events & training', "showPage('trade',3)");
  if (/event|exhibition|fse|show/.test(t))             return jump('Open Trade, events & training', "showPage('trade',3)");
  if (/showroom|visit|book/.test(t))                   return jump('Book the showroom', 'openBookingModal()');
  if (/cert|kitemark|conformity|med|mer|nta/.test(t))  return jump('Open the Product portal', 'openProductPortal()');
  return '';
}

function emberOpen(i) {
  const f = (EMBER.lastDocs || [])[i];
  if (f && typeof openDocFile === 'function') openDocFile(f);
}

// ── Keyboard ─────────────────────────────────────────────────
// Ctrl/⌘+K opens Ember from anywhere; Escape closes it — unless the
// reader is open, in which case that takes the Escape first.
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === 'k') {
    e.preventDefault();
    toggleEmber(true);
    return;
  }
  if (e.key === 'Escape' && EMBER.open) {
    const reader = document.getElementById('doc-modal');
    if (reader && !reader.classList.contains('hidden')) return;
    toggleEmber(false);
  }
});
