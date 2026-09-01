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

// 1 Sep 2026 — "we'd like to change the name to Josh 2.0". One place.
// The class names stay ember-* (see index.html for why); everything a
// person reads comes from here.
function _emberName()    { return _emberCfg().name || 'Josh 2.0'; }
function _emberShort()   { return _emberName().split(/[\s.]/)[0] || _emberName(); }
function _emberInitial() { return (_emberName().trim()[0] || 'J').toUpperCase(); }

// "Welcome back, Bodhi!" — the Jotform agent greets you by name when it
// knows it, and the hub always does: MSAL has already told us who is
// signed in. Falls back to no name rather than to a placeholder.
function _emberWho() {
  const a = (window.AUTH && window.AUTH.account) || {};
  const nm = String(a.displayName || a.name || '').trim();
  if (!nm) return '';
  const first = nm.split(/[\s,]+/)[0];
  return (first && first.length > 1) ? first : '';
}

// Paint the name into the shell once, so the markup never has to hold
// it twice.
function _emberBrand() {
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  set('ember-name', _emberName());
  set('ember-av', _emberInitial());
  set('ember-launch-label', 'Ask ' + _emberShort());
  // The Jotform Josh 2.0 marketing built carries a role under the name —
  // "Product Specialist" — and it does a lot of work: you know what to
  // ask before you've typed anything.
  set('ember-tagline', _emberCfg().role || _emberCfg().tagline || 'CheckFire\u2019s assistant');
  const inp = document.getElementById('ember-input');
  if (inp) inp.placeholder = 'Ask ' + _emberShort() + ' anything…';
  const launch = document.querySelector('.ember-launch');
  if (launch) launch.setAttribute('aria-label', 'Ask ' + _emberName());
}
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
    if (!EMBER.booted) { EMBER.booted = true; _emberBrand(); _emberBoot(); }
    setTimeout(() => { const i = document.getElementById('ember-input'); if (i) i.focus(); }, 260);
  }
}

function _emberBoot() {
  const body = document.getElementById('ember-body');
  if (!body) return;

  if (_emberMode() === 'copilot') {
    body.innerHTML = `<iframe class="ember-frame" src="${escAttr(safeUrl(_emberCfg().copilotEmbedUrl, ''))}"
      title="${escAttr(_emberName())}" frameborder="0"
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
        <h3>${escHtml(_emberWho() ? 'Hi ' + _emberWho() + ' — I’m ' + _emberName() : 'Hi, I’m ' + _emberName())}</h3>
        <p>${mode === 'claude'
            ? escHtml(_emberCfg().role || 'Product specialist') + '. Ask me about a product, the paperwork behind it, a campaign or a training date — or give me something to write. How can I help?'
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

  const av = escHtml(_emberInitial());
  body.innerHTML = EMBER.view.map(m => `
    <div class="ember-msg ${m.role === 'user' ? 'you' : 'ember'}">
      ${m.role === 'user' ? '' : `<span class="ember-av">${av}</span>`}
      <div class="ember-bubble">${m.html}</div>
    </div>`).join('') +
    // The live bubble the streamed answer lands in, word by word.
    (EMBER.busy ? `<div class="ember-msg ember"><span class="ember-av">${av}</span>
      <div class="ember-bubble" id="ember-live"><span class="ember-dots"><i></i><i></i><i></i></span>
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

// Turn a thrown error into something a person can act on. Dev tools are
// disabled by policy on the CheckFire build, so "Failed to fetch" in the
// panel was the end of the road — there was nowhere to look for more.
// This says what the hub actually tried and what it means.
function _emberError(e) {
  const msg = String((e && e.message) || e || 'Unknown error');
  const url = _emberAiUrl();

  // fetch() throws a TypeError for anything that stopped the request
  // reaching the server OR stopped the browser accepting the reply:
  // no network, DNS, a web filter, or a failed CORS check.
  const network = (e instanceof TypeError) ||
                  /failed to fetch|networkerror|load failed/i.test(msg);

  if (network) {
    return `<p>I couldn’t reach the assistant service, so there’s nothing to answer with yet.</p>
      <p>The hub asked for <code>${escHtml(url + '/chat')}</code> and the browser refused
      the request before any reply came back. That is almost always one of three things:
      the service is asleep or down, the network is blocking that address, or the service
      answered without the permission header this site needs.</p>
      <p>Nothing is wrong with what you typed — try again in a moment, and if it keeps
      happening send this whole message on.</p>`;
  }

  return `<p>Something went wrong there — ${escHtml(msg)}.</p>
    <p>Try again in a moment. If it keeps happening, send this message on —
    it names the actual fault.</p>`;
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
    html = _emberError(e);
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

// ── Deciding whether this question needs a document at all ───
//
// 1 Sep 2026 — David: "we need to ensure it's working like an actually
// AI chat bot. Like talking to you."
//
// Claude was already behind it. What made it feel like a search box was
// that EVERY message ran a full SharePoint search first — three sites,
// three queries each, then up to eight documents downloaded and read —
// before it would answer "morning" or "make that shorter". Five or six
// seconds of "looking through SharePoint…" for a reply that never
// needed a single file.
//
// So it now decides. Anything that reads like a document question still
// gets the full retrieval; conversation, follow-ups and writing tasks
// go straight to Claude with the hub summary and come back at once. The
// summary of launches, campaigns and training is sent either way — it
// is already in the cache and costs nothing.
const EMBER_CHATTY = /^(hi|hey|hello|morning|afternoon|evening|thanks|thank you|cheers|ta|ok|okay|great|perfect|nice|yes|no|please do|go on)\b/i;
const EMBER_FOLLOWUP = /^(make it|make that|shorter|longer|again|do it again|rewrite|reword|try again|expand|more like|less|and |also |what about|explain that|why\b|how so|really\?|carry on|continue)/i;
const EMBER_WRITING = /\b(write|draft|rewrite|reword|shorten|tighten|summarise|summarize|caption|subject line|headline|strapline|post|tweet|idea|ideas|brainstorm|suggest|tone)\b/i;
const EMBER_DOCWORD = /\b(datasheet|data sheet|certificate|certification|declaration|conformity|msds|sds|pif|kitemark|med\b|mer\b|nta|manual|instruction|document|spec|specification|pdf|policy|guideline|brochure|price|part number|code)\b/i;

// Answers are short by default. They stop being short when someone asks
// for something that is genuinely long — a draft, a list, a full
// explanation — and then the cap lifts rather than the answer stopping
// mid-sentence.
const EMBER_LONG = /\b(draft|write|rewrite|email|post|article|blog|script|list|bullet|step by step|in detail|in full|explain|summar|brief|proposal|agenda)\b/i;

function _emberTokens(q) {
  const cfg = _emberCfg();
  return EMBER_LONG.test(String(q || ''))
    ? (cfg.maxTokensLong || 2000)
    : (cfg.maxTokens || 700);
}

function _emberNeedsDocs(q) {
  if (_emberCfg().skipSearchWhenChatty === false) return true;
  const t = String(q || '').trim();
  if (!t) return false;
  if (EMBER_DOCWORD.test(t)) return true;          // always search for these
  if (t.length < 14) return false;                 // "hi", "thanks", "go on"
  if (EMBER_CHATTY.test(t)) return false;
  if (EMBER.turns.length > 1 && EMBER_FOLLOWUP.test(t)) return false;
  if (EMBER_WRITING.test(t)) return false;         // a writing job, not a lookup
  return true;
}

// ── Claude mode ──────────────────────────────────────────────

const EMBER_META = '\n<<<META>>>';

// ── Tap-to-ask follow-ups ────────────────────────────────────
//
// 1 Sep 2026 — David pointed at the Jotform agent marketing had already
// built (eu.jotform.com, "Josh 2.0 · Product Specialist") and said: chat
// like that. Two things make that agent feel like a conversation rather
// than a query box, and neither is the model:
//
//   1. It answers in ONE line and then offers the next thing it can do
//      — "…if you want, I can also help you choose the right
//      extinguisher for home, office, or vehicle use."
//   2. Every message comes with two tappable options, so you are never
//      staring at an empty box wondering what it can do.
//
// The offer is a prompt instruction (see the Function app). The taps are
// this: the assistant ends its reply with one line —
//
//     SUGGEST: Show me the datasheet | Draft a customer email
//
// — which never reaches the screen. It is stripped out here, turned into
// buttons, and stripped from the stored turn as well so it can't leak
// into the next answer.
const EMBER_SUGGEST = /\n[ \t]*SUGGEST[ \t]*:[ \t]*(.*)$/i;

function _emberSplitSuggest(text) {
  const m = EMBER_SUGGEST.exec(String(text || ''));
  if (!m) return { text: String(text || '').trim(), chips: [] };
  const chips = m[1].split('|').map(x => x.trim())
    .filter(x => x && x.length < 70).slice(0, 3);
  return { text: String(text).slice(0, m.index).trim(), chips };
}

// Mid-stream the SUGGEST line arrives a character at a time, so hide any
// trailing fragment of it too — otherwise you watch "SUGG" appear under
// the answer and then vanish.
const EMBER_SUGGEST_PARTIAL = /\n[ \t]*S(U(G(G(E(S(T[ \t]*:?[^\n]*)?)?)?)?)?)?$/i;

function _emberVisible(text) {
  return String(text || '').replace(EMBER_SUGGEST, '').replace(EMBER_SUGGEST_PARTIAL, '');
}

// If the assistant didn't write a SUGGEST line — an older Function app
// that predates the contract, or a turn where it simply didn't — the
// panel should still offer somewhere to go. The Jotform agent never
// leaves you with a dead end, and neither should this.
function _emberFallbackChips(q) {
  const t = String(q || '').toLowerCase();
  if (/cert|kitemark|conformity|declaration|msds|sds|pif|\bmed\b|\bmer\b|nta/.test(t))
    return ['Is there a newer version?', 'Show me the datasheet'];
  if (/datasheet|data sheet|dimension|spec|rating|capacity|weight/.test(t))
    return ['Show me the datasheet', 'What else is in that range?'];
  if (/launch|npd|new product/.test(t))
    return ['What’s launching next?', 'Show me the launch pack'];
  // Intent before subject: "draft a LinkedIn post" is a writing job that
  // happens to contain the word "post", not a campaign question.
  if (/write|draft|rewrite|reword|shorten|caption|headline|version/.test(t))
    return ['Make it shorter', 'Give me another version'];
  if (/campaign|social|email|post|blog/.test(t))
    return ['What assets do we have?', 'Draft a social post'];
  if (/training|course|session/.test(t))
    return ['When is the next session?', 'Book me on'];
  if (/event|exhibition|fse|show|stand/.test(t))
    return ['What’s in the event pack?', 'When is it?'];
  return ['Tell me more', 'Where did that come from?'];
}

function _emberFollowups(chips) {
  if (!chips || !chips.length) return '';
  return `<div class="ember-next">${chips.map(c =>
    `<button class="ember-chip" onclick="emberAsk(${JSON.stringify(c).replace(/"/g, '&quot;')})">${escHtml(c)}</button>`
  ).join('')}</div>`;
}

async function _emberClaude(question) {
  if (!(window.AUTH && window.AUTH.account)) {
    return `<p>Sign in with your CheckFire account first and I can read our documents for you.</p>`;
  }

  const cfg = _emberCfg();
  let found = [];
  const docs = [];

  if (_emberNeedsDocs(question)) {
    _emberStatus('looking through SharePoint…');
    found = await _emberFind(question);

    // Only documents worth reading — no images or videos.
    const readable = found.filter(f => EMBER_READABLE.test(f.name || '')).slice(0, cfg.maxDocs || 8);
    _emberStatus(readable.length ? `reading ${readable.length} document${readable.length === 1 ? '' : 's'}…` : 'thinking…');

    // A downloadUrl is short-lived and pre-authenticated: it is a link to
    // ONE file the signed-in user can already open, and it is the only
    // thing about that file that leaves the browser.
    await Promise.all(readable.map(async f => {
      try {
        const meta = await graphFetch(
          `/drives/${f._driveId}/items/${f.id}?$select=id,name,size,@microsoft.graph.downloadUrl`);
        const url = meta && meta['@microsoft.graph.downloadUrl'];
        if (!url) return;
        docs.push({ name: f.name, site: f._site, url, size: f.size });
      } catch (_) { /* skip it */ }
    }));
  } else {
    console.info('[Josh] conversational turn — answering without a document search.');
  }

  const hubContext = await _emberHubContext();
  EMBER.lastDocs = found.slice(0, 12);

  _emberStatus('thinking…');

  const wantStream = cfg.stream !== false;

  // Content-Type is 'text/plain', NOT 'application/json', and that is
  // deliberate — do not "tidy" it back.
  //
  // 27 Aug 2026. An application/json POST is not a CORS-"simple" request,
  // so Chrome sends an OPTIONS preflight first. The Function app answers
  // that preflight **204 with no CORS headers at all** (only Date and
  // Server survive — Azure Functions drops custom headers on a 204), so
  // Chrome rejects it and `fetch` throws "Failed to fetch" without ever
  // sending the real request. The POST itself was always fine.
  //
  // text/plain IS on the CORS-safelist, so no preflight is sent, the
  // POST goes straight out, and its own header satisfies the browser.
  const res = await fetch(_emberAiUrl() + '/chat', {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
    body: JSON.stringify({
      messages: EMBER.turns.slice(-12),
      docs,
      context: hubContext,
      fileList: found.slice(0, 20).map(f => ({ name: f.name, site: f._site })),
      maxCharsPerDoc: cfg.maxCharsPerDoc || 6000,
      // Ignored by a Function app that predates 1 Sep 2026, which is
      // why the hub still works against the old one — it just answers
      // as Ember, all at once, instead of as Josh 2.0, word by word.
      persona: _emberName(),
      role: cfg.role || '',
      who: _emberWho(),
      maxTokens: _emberTokens(question),
      stream: wantStream,
    }),
  });

  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    if (res.status === 404) throw new Error('the AI proxy isn’t deployed yet');
    throw new Error(`the AI proxy returned ${res.status}${detail ? ' — ' + detail.slice(0, 160) : ''}`);
  }

  const ctype = String(res.headers.get('content-type') || '');
  const streaming = wantStream && res.body && typeof res.body.getReader === 'function'
                    && !/application\/json/i.test(ctype);

  let answer = '', usedNames = null;

  if (streaming) {
    const out = await _emberStreamIn(res);
    answer = out.text;
    usedNames = out.used;
  } else {
    const data = await res.json();
    answer = String(data.text || data.answer || '').trim();
    usedNames = Array.isArray(data.used) ? data.used.map(String) : null;
  }

  if (!answer) throw new Error('the AI proxy sent an empty answer');

  const split = _emberSplitSuggest(answer);
  answer = split.text;
  if (!answer) throw new Error('the AI proxy sent an empty answer');

  // The stored turn is the visible answer only — the SUGGEST line is
  // scaffolding, and feeding it back would teach it to write more.
  EMBER.turns.push({ role: 'assistant', text: answer });

  // Only show source cards for documents it actually used, when the
  // proxy tells us; otherwise the best few it was given.
  const sources = usedNames
    ? EMBER.lastDocs.filter(f => usedNames.some(n => String(f.name).toLowerCase() === n.toLowerCase()))
    : EMBER.lastDocs.slice(0, 4);

  const chips = split.chips.length ? split.chips : _emberFallbackChips(question);

  return _emberMarkdown(answer) + _emberFollowups(chips)
       + _emberSources(sources) + _emberShortcut(question.toLowerCase());
}

// Paint the answer as it arrives. This is the single biggest reason the
// old panel felt like a machine rather than a conversation: you asked,
// nothing happened for eight seconds, then a wall of text appeared.
//
// The body is plain text. Anything the server wants to tell us
// afterwards — which documents it used — comes after a sentinel line at
// the very end, so the stream itself stays readable.
async function _emberStreamIn(res) {
  const body   = document.getElementById('ember-body');
  const reader = res.body.getReader();
  const dec    = new TextDecoder();
  let raw = '';

  const paint = () => {
    const live = document.getElementById('ember-live');
    if (!live) return;
    const cut = raw.indexOf(EMBER_META);
    const shown = _emberVisible(cut < 0 ? raw : raw.slice(0, cut));
    if (!shown) return;
    live.innerHTML = _emberMarkdown(shown) + '<span class="ember-caret"></span>';
    if (body) body.scrollTop = body.scrollHeight;
  };

  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      raw += dec.decode(value, { stream: true });
      paint();
    }
    raw += dec.decode();
  } catch (e) {
    // A stream that dies half way still leaves a usable answer on
    // screen — better than swapping it for an error.
    console.info('[Josh] stream ended early:', e.message);
  }

  const cut = raw.indexOf(EMBER_META);
  const text = (cut < 0 ? raw : raw.slice(0, cut)).trim();
  let used = null;
  if (cut >= 0) {
    try { used = (JSON.parse(raw.slice(cut + EMBER_META.length)).used || []).map(String); }
    catch (_) { /* the answer matters, the footnote doesn't */ }
  }
  return { text, used };
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
    // The reader is a page now, not an overlay, so Ember can never be
    // covering it — but if both are somehow up, the reader takes Escape.
    const reader = document.getElementById('page-reader');
    if (reader && reader.classList.contains('active')) return;
    toggleEmber(false);
  }
});
