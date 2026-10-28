/**
 * CheckFire Marketing Hub — Sales Spotlight (Product Portal)
 *
 * 28 Oct 2026. A band above "What are you looking for?" with this
 * fortnight's topic and its top three, opening a page with the full top
 * ten, the month-by-month line and the previous spotlights.
 *
 * Figures: the sales dashboard's customer-items.json on the Sales &
 * Marketing site (see HUB_CONFIG.spotlight). Measured NetSuite sales and
 * quantities, per customer, per product class, per item, for each month
 * of last year and this year:
 *   {year, closed, names, c: {COMPANY: {class: {item: {a:[24], q:[24]}}}}}
 *   index 0-11 = Jan-Dec last year, 12-23 = Jan-Dec this year;
 *   `closed` months of this year are complete, the next one is to date.
 *
 * Topic: the "Spotlight" list on the Product Portal site, or
 * HUB_CONFIG.spotlight.fallback until that list exists.
 *
 * Depends on graph.js (graphFetch, resolveSiteId, _fetchPortalList,
 * _ppField, _ppTruthy, fmtSpDate) and ui.js (escHtml, escAttr, showPage).
 */

// `denied`: the figures couldn't be read this visit (no access to the
// Sales & Marketing site, or no file), so the band stops asking.
// `tok`: the newest band render; an older one finishing late must not
// overwrite it (the portal can ask for the band twice on first open).
const SS = { topics: [], cur: -1, data: null, file: null, loading: null, denied: false, tok: 0 };

const SS_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function _ssSlug(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function _ssGbp(n)  { return '£' + Math.round(n).toLocaleString('en-GB'); }
function _ssInt(n)  { return Math.round(n).toLocaleString('en-GB'); }

// A SharePoint date-only column arrives as UTC midnight of the local
// day (in BST, 27 Oct is "2026-10-26T23:00:00Z"), so it is read as a
// local date, never sliced as a string.
function _ssDay(v) {
  if (!v) return null;
  const d = new Date(v);
  if (isNaN(d)) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function _ssRange(t) {
  const fmt = d => d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });
  if (t.start && t.end) return `${fmt(t.start)} – ${fmt(t.end)}`;
  if (t.start)          return `From ${fmt(t.start)}`;
  return '';
}

// ── Topics ────────────────────────────────────────────────────
// The current topic is the latest one that has started and not ended.
// If the fortnight has run out and nobody has added the next row yet,
// the latest started one stays up rather than the band vanishing from
// the meeting screen. Rows starting in the future wait their turn.
async function _ssTopics() {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let rows = [];
  try { rows = await _fetchPortalList(HUB_CONFIG.lists.spotlight); }
  catch (e) { console.info('[Spotlight] list not read —', e.message); }

  let topics = (rows || [])
    .filter(r => r && r.Title && _ppTruthy(_ppField(r, ['Show', 'Active', 'Visible'])))
    .map(r => ({
      title:    String(r.Title).trim(),
      type:     /item|product/i.test(String(_ppField(r, ['Type', 'Kind']) || '')) ? 'items' : 'customers',
      category: String(_ppField(r, ['Category', 'Class', 'ProductGroup', 'Product Group']) || '').trim(),
      items:    String(_ppField(r, ['Items', 'ItemCodes', 'Item Codes', 'Codes']) || '').trim(),
      tag:      String(_ppField(r, ['Tag', 'Label', 'Range']) || '').trim(),
      start:    _ssDay(_ppField(r, ['Start', 'StartDate', 'Start Date', 'From'])),
      end:      _ssDay(_ppField(r, ['End', 'EndDate', 'End Date', 'To'])),
    }))
    .filter(t => !t.start || t.start <= today)
    .sort((a, b) => (b.start ? b.start.getTime() : 0) - (a.start ? a.start.getTime() : 0));

  // The config topic stands in only while there is no list at all. A
  // list whose rows are all hidden or not started yet means "no band",
  // so the product team can take it down themselves.
  if (!(rows || []).length) {
    topics = ((HUB_CONFIG.spotlight && HUB_CONFIG.spotlight.fallback) || [])
      .map(t => ({ ...t, start: null, end: null }));
  }

  const live = topics.findIndex(t => !t.end || t.end >= today);
  SS.cur = live >= 0 ? live : (topics.length ? 0 : -1);
  SS.topics = topics;
  return topics;
}

// ── Data ──────────────────────────────────────────────────────
// About 8 MB, so it is fetched once per visit and kept in memory. The
// download URL is pre-authenticated and short-lived; asking for it with
// the item means one Graph call plus one plain download.
function _ssLoad() {
  if (SS.data) return Promise.resolve(SS.data);
  if (SS.loading) return SS.loading;
  const cfg = HUB_CONFIG.spotlight || {};
  SS.loading = (async () => {
    const siteId = await resolveSiteId(cfg.dataSite);
    const path   = cfg.dataPath.split('/').map(encodeURIComponent).join('/');
    // No $select: on a single item fetched by path, Graph drops
    // @microsoft.graph.downloadUrl when $select is given (seen 28 Oct on
    // the preview), even when it is named. The V2 dashboard reads the
    // same file the same way.
    const item   = await graphFetch(`/sites/${siteId}/drive/root:/${path}`);
    const url    = item['@microsoft.graph.downloadUrl'];
    if (!url) throw new Error('no download URL for ' + cfg.dataPath);
    const res = await fetch(url);
    if (!res.ok) throw new Error('download returned ' + res.status);
    SS.data = await res.json();
    SS.file = { modified: item.lastModifiedDateTime };
    return SS.data;
  })();
  SS.loading.catch(() => { SS.loading = null; });
  return SS.loading;
}

// "SD-CFP-*; FP-COM-001-RED" → a test on an item code. A trailing *
// matches by prefix, so a range with pallet and kit codes is one entry.
function _ssItemTest(spec) {
  const parts = String(spec || '').split(/[;,\n]/).map(s => s.trim().toUpperCase()).filter(Boolean);
  if (!parts.length) return null;
  return code => {
    const c = String(code).toUpperCase();
    return parts.some(p => p.endsWith('*') ? c.startsWith(p.slice(0, -1)) : c === p);
  };
}

// Category and Items both narrow: a line counts when its class is one of
// the categories (if any are given) AND its code is one of the items (if
// any are given). Neither given matches nothing, deliberately.
//
// Like-for-like: the COMPLETED months of this year against the same
// months of last year (Jan–Sep v Jan–Sep), never the month in progress,
// so a change on screen is never a part month against a whole one.
function _ssCompute(t, data) {
  const cats    = String(t.category || '').split(/[;,]/).map(_ssSlug).filter(Boolean);
  const itemOk  = _ssItemTest(t.items);
  const closed  = Math.max(0, Math.min(12, Number(data.closed) || 0));
  const out     = { rows: [], months: new Array(12).fill(0), lyMonths: new Array(12).fill(0),
                    total: 0, qty: 0, lastYear: 0, likeTy: 0, likeLy: 0, closed, buyers: 0 };
  if (!cats.length && !itemOk) return out;

  const by = new Map();
  for (const [comp, classes] of Object.entries(data.c || {})) {
    for (const [cls, items] of Object.entries(classes)) {
      if (cats.length && cats.indexOf(_ssSlug(cls)) < 0) continue;
      for (const [code, v] of Object.entries(items)) {
        if (itemOk && !itemOk(code)) continue;
        const key = t.type === 'items' ? code : comp;
        const e = by.get(key) || { key, a: 0, q: 0, ty: 0, ly: 0, lyAll: 0 };
        for (let m = 0; m < 12; m++) {
          const ta = v.a[12 + m] || 0, la = v.a[m] || 0;
          e.a += ta; e.q += (v.q && v.q[12 + m]) || 0; e.lyAll += la;
          out.months[m] += ta; out.lyMonths[m] += la;
          if (m < closed) { e.ty += ta; e.ly += la; }
        }
        by.set(key, e);
      }
    }
  }
  for (const e of by.values()) {
    out.total += e.a; out.qty += e.q; out.lastYear += e.lyAll;
    out.likeTy += e.ty; out.likeLy += e.ly;
  }
  // Customers with net sales this year (credits can take one to nil).
  out.buyers = [...by.values()].filter(e => e.a > 0).length;
  out.rows = [...by.values()]
    .filter(e => e.a > 0)
    .sort((x, y) => y.a - x.a)
    .map(e => ({ ...e, name: t.type === 'items' ? ((data.names || {})[e.key] || e.key) : e.key }));
  return out;
}

// Words for a topic. One picked by item codes alone has no category to
// name, so it reads "sales of these products" rather than a gap.
function _ssWords(t) {
  const c = t.category;
  return {
    sales:   c ? `${c} sales` : 'Sales of these products',
    salesLc: c ? `${c} sales` : 'sales of these products',
    adj:     c ? `${c} ` : '',
    chart:   c ? `${c} by month` : 'By month',
  };
}

// The months this year that have any figures in them: January up to the
// month in progress.
function _ssMonthCount(data) { return Math.min(12, (Number(data.closed) || 0) + 1); }

// "Jan–Sep" for the like-for-like window.
function _ssWindow(closed) { return closed <= 1 ? SS_MONTHS[0] : `${SS_MONTHS[0]}–${SS_MONTHS[closed - 1]}`; }

// A change, with its sign and arrow, never colour alone (brand rule:
// red only for a fall, green for a rise, grey for nothing to compare).
function _ssChange(ty, ly, opts) {
  const o = opts || {};
  if (!(ly > 0)) return ty > 0 ? `<span class="ss-chg new">New${o.long ? ' this year' : ''}</span>` : '';
  const p = (ty - ly) / Math.abs(ly) * 100;
  if (Math.abs(p) < 0.5) return `<span class="ss-chg flat">No change</span>`;
  const up = p > 0;
  return `<span class="ss-chg ${up ? 'up' : 'down'}"><i aria-hidden="true">${up ? '▲' : '▼'}</i>${up ? '+' : '−'}${Math.abs(p).toFixed(0)}%</span>`;
}

function _ssShare(a, total) { return total > 0 ? a / total * 100 : 0; }
function _ssPct(n) { return n >= 10 ? n.toFixed(0) + '%' : n.toFixed(1) + '%'; }

// The mark on a spotlight card: the category, or the initials of the
// product words in the title ("Who is buying LFX fire blankets" → LFB).
function _ssTopicMark(x) {
  if (x.category) return x.category.slice(0, 3).toUpperCase();
  const w = String(x.title).replace(/^(who('s| is) buying|who buys|which)\s+/i, '')
    .split(/[\s:-]+/).filter(t => t && !/^(the|new|a|an|of|and|which|models|sell|this|year)$/i.test(t));
  return w.slice(0, 3).map(t => /^\d/.test(t) ? t.replace(/[^0-9a-z]/gi, '').slice(0, 2) : t[0]).join('').toUpperCase() || 'S';
}

// Initials for a customer or product chip: "O. HEAP & SON" → "OH".
function _ssInitials(name) {
  const w = String(name).replace(/[^A-Za-z0-9 ]/g, ' ').split(/\s+/)
    .filter(x => x && !/^(LTD|LIMITED|THE|AND|OF|UK|PLC|WH|CO)$/i.test(x));
  return ((w[0] || '?')[0] + ((w[1] || '')[0] || '')).toUpperCase();
}

// Counts a figure up from zero once. A figure inside a section still
// waiting to scroll into view is left for _ssReveal to start later.
// data-ss-fmt: £ (default), 'int' or 'pct'.
// Skipped entirely when the viewer asks for reduced motion.
function _ssCountUp(root) {
  const els = (root || document).querySelectorAll('[data-ss-count]');
  const still = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  els.forEach(el => {
    if (el.closest('.ss-reveal:not(.ss-seen)')) return;
    const to = Number(el.getAttribute('data-ss-count')) || 0;
    const fmt = el.getAttribute('data-ss-fmt');
    const show = v => fmt === 'pct' ? _ssPct(v) : fmt === 'int' ? _ssInt(v) : _ssGbp(v);
    el.removeAttribute('data-ss-count');
    if (still || !to) return;
    const t0 = performance.now(), dur = 900;
    const step = now => {
      const k = Math.min(1, (now - t0) / dur), e = 1 - Math.pow(1 - k, 3);
      el.textContent = show(to * e);
      if (k < 1) requestAnimationFrame(step); else el.textContent = show(to);
    };
    requestAnimationFrame(step);
  });
}

// Twelve small bars: this year to date. Complete months in the light
// tone, the month in progress hatched, months to come left empty.
function _ssSpark(r, nm) {
  const max = Math.max(1, ...r.months.slice(0, nm), ...r.lyMonths);
  return `<div class="ss-spark" aria-hidden="true">${r.months.map((v, m) => `
    <span class="${m === nm - 1 ? 'part' : m >= nm ? 'future' : ''}"><i style="--h:${m >= nm ? 0 : Math.max(2, v / max * 100).toFixed(1)}%"></i></span>`).join('')}</div>
    <div class="ss-spark-lab" aria-hidden="true">${SS_MONTHS.map(x => `<span>${x[0]}</span>`).join('')}</div>`;
}

// ── The band on the Product Portal front ──────────────────────
async function renderSpotlightBand() {
  const host = document.getElementById('pp-spotlight');
  if (!host) return;
  if (typeof getAccessToken !== 'function' || window.HUB_DEMO_MODE || SS.denied) { host.innerHTML = ''; return; }
  const tok = ++SS.tok;

  const topics = await _ssTopics();
  if (tok !== SS.tok) return;
  if (SS.cur < 0) { host.innerHTML = ''; return; }
  const t = topics[SS.cur];

  // Skeleton only while the figures are actually on their way; with
  // them already in memory the band is redrawn in one go, no flicker.
  if (!SS.data) {
    host.innerHTML = `
      <div class="pp-band ss-band">
        <div class="pp-band-head">
          <h2 class="pp-band-title">Sales Spotlight</h2>
          <span class="pp-band-note">A new topic every fortnight</span>
        </div>
        <div class="ss-feature-skel skeleton"></div>
      </div>`;
  }

  let data;
  try { data = await _ssLoad(); }
  catch (e) {
    // No access to the Sales & Marketing site, or the file isn't there.
    // Sales figures are not for everyone, so this is a quiet absence,
    // and it isn't asked again until the page is reloaded.
    console.info('[Spotlight] figures not read —', e.message);
    SS.denied = true;
    if (tok === SS.tok) host.innerHTML = '';
    return;
  }
  if (tok !== SS.tok) return;

  const r     = _ssCompute(t, data);
  const top3  = r.rows.slice(0, 3);
  const w     = _ssWords(t);
  const range = _ssRange(t);
  const nm    = _ssMonthCount(data);
  const topN  = (HUB_CONFIG.spotlight || {}).topN || 10;
  const t3    = top3.reduce((s, x) => s + x.a, 0);
  const max3  = top3.length ? top3[0].a : 1;

  host.innerHTML = `
   <div class="pp-band ss-band">
    <div class="pp-band-head">
      <h2 class="pp-band-title">Sales Spotlight</h2>
      <span class="pp-band-note">A new topic every fortnight</span>
    </div>
    <article class="ss-feature" role="button" tabindex="0"
             onclick="openSpotlight(${SS.cur})" onkeydown="if(event.key==='Enter')openSpotlight(${SS.cur})">
      <div class="ss-feature-fig">
        <div class="ss-feature-top">
          <span class="ss-live"><span class="ss-live-dot"></span>${range ? escHtml(range) : 'Live'}</span>
          <span class="ss-feature-eyebrow">${escHtml(w.sales)} · ${data.year} to date</span>
        </div>
        <b class="ss-big" data-ss-count="${Math.round(r.total)}">${_ssGbp(r.total)}</b>
        ${r.closed && r.likeLy ? `<div class="ss-feature-chg">${_ssChange(r.likeTy, r.likeLy, { long: true })}<span>${_ssWindow(r.closed)} v ${data.year - 1}</span></div>` : ''}
        ${_ssSpark(r, nm)}
      </div>
      <div class="ss-feature-body">
        <div class="px-card-eyebrow">This fortnight${t.category ? ' · ' + escHtml(t.category) : ''}</div>
        <h3 class="ss-feature-title">${escHtml(t.title)}</h3>
        ${top3.length ? `<ol class="ss-top3">${top3.map((x, i) => `
          <li style="--i:${i}">
            <span class="ss-rk">${i + 1}</span>
            <span class="ss-av">${escHtml(_ssInitials(x.name))}</span>
            <span class="ss-nm"><b>${escHtml(x.name)}</b><span class="ss-bar"><i style="--w:${(x.a / max3 * 100).toFixed(1)}%"></i></span></span>
            <span class="ss-val">${_ssGbp(x.a)}<small>${_ssPct(_ssShare(x.a, r.total))} of sales</small></span>
          </li>`).join('')}</ol>
          <p class="ss-feature-foot">These three bought <b>${_ssPct(_ssShare(t3, r.total))}</b> of ${escHtml(w.salesLc)} this year, out of ${_ssInt(t.type === 'items' ? r.rows.length : r.buyers)} ${t.type === 'items' ? 'products' : 'customers'}.</p>`
          : `<p class="px-card-meta">No sales found this year. Check the Category or Items on the Spotlight list.</p>`}
        <span class="ss-go">See the full top ${topN}
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
        </span>
      </div>
    </article>
   </div>`;
  _ssCountUp(host);
  if (!r.rows.length) console.warn(`[Spotlight] "${t.title}" matched no sales — Category "${t.category}", Items "${t.items}".`);
}

// ── The Spotlight page ────────────────────────────────────────
// 28 Oct 2026, David: "when you click in to it … impress me". So the
// page opens as an event: a beam sweeps the dark hero and then follows
// the pointer, the title rises word by word, the top-ten ring draws,
// the podium grows from the floor, and every figure counts up as it
// comes into view. The figures are exactly the ones on the card.
async function openSpotlight(i) {
  const host = document.getElementById('spotlight-body');
  if (!host) return;
  showPage('spotlight');
  if (typeof updateNavActive === 'function') updateNavActive('portal');

  let data;
  try {
    if (!SS.topics.length) await _ssTopics();
    data = await _ssLoad();
  } catch (e) {
    host.innerHTML = `<div class="px-empty"><h3>Figures not available</h3><p>The sales figures couldn't be read: ${escHtml(e.message)}</p></div>`;
    return;
  }
  const t = SS.topics[i];
  if (!t) { openProductPortal(); return; }
  SS.open = i;

  const topN  = (HUB_CONFIG.spotlight || {}).topN || 10;
  const r     = _ssCompute(t, data);
  const rows  = r.rows.slice(0, topN);
  const pod   = rows.slice(0, 3);
  const rest  = rows.slice(3);
  const max   = rows.length ? rows[0].a : 1;
  const nm    = _ssMonthCount(data);
  const w     = _ssWords(t);
  const yr    = data.year;
  const range = _ssRange(t);
  const isNow = i === SS.cur;
  const noun  = t.type === 'items' ? 'products' : 'customers';
  const count = t.type === 'items' ? r.rows.length : r.buyers;
  const topSum = rows.reduce((s, x) => s + x.a, 0);
  const others = Math.max(0, r.total - topSum);
  const share  = _ssShare(topSum, r.total);
  const win   = _ssWindow(r.closed);
  const cmp   = r.closed > 0;
  const qty   = !!data.hasQty;
  const items = t.type === 'items';

  // Month chart: this year (ink) beside last year (grey), same scale.
  const cMax = Math.max(1, ...r.months.slice(0, nm), ...r.lyMonths);
  const peak = r.months.slice(0, r.closed).reduce((b, v, m) => v > b.v ? { v, m } : b, { v: -Infinity, m: -1 });

  // Podium order on screen: 2, 1, 3 — the winner in the middle, raised.
  const podOrder = [1, 0, 2].filter(n => pod[n]);

  // Ring: the top ten's share of all sales, drawn as an arc.
  const R = 54, C = 2 * Math.PI * R;

  const prev = SS.topics.map((x, n) => ({ x, n })).filter(o => o.n !== i);
  SS.view = { t, r, data, nm, yr, mode: 'cum' };

  // Title words rise in one after another.
  const words = String(t.title).split(/\s+/).filter(Boolean)
    .map((wd, k) => `<span class="ss-w" style="--k:${k}">${escHtml(wd)}</span>`).join(' ');

  host.innerHTML = `
    <section class="ss-stage" id="ss-stage">
      <div class="ss-beam" aria-hidden="true"></div>
      <div class="ss-stage-in">
        <div class="ss-stage-copy">
          <div class="ss-kicker"><span class="ss-live-dot"></span>Sales Spotlight${range ? ' · ' + escHtml(range) : ''}${isNow ? '' : ' · previous'}</div>
          <h1 class="ss-stage-title">${words}</h1>
          <p class="ss-stage-sub">Our top ${topN} ${noun} by ${escHtml(w.salesLc)}, 1 January to date${isNow ? '' : ", with today's figures"}.</p>
          <div class="ss-stage-actions">
            <button class="ss-present-btn" onclick="ssPresent()">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>
              Present
            </button>
            ${rows.length ? `<a class="ss-jump" href="#ss-podium" onclick="document.getElementById('ss-podium').scrollIntoView({behavior:'smooth',block:'start'});return false">See the podium ↓</a>` : ''}
          </div>
        </div>
        <div class="ss-stage-fig">
          <span class="ss-stage-label">${escHtml(w.sales)} · ${yr} to date</span>
          <b class="ss-big xl" data-ss-count="${Math.round(r.total)}">${_ssGbp(r.total)}</b>
          ${cmp && r.likeLy ? `<div class="ss-stage-chg">${_ssChange(r.likeTy, r.likeLy, { long: true })}<span>${win} ${yr} v ${win} ${yr - 1}</span></div>` : ''}
          <div class="ss-stage-stats">
            <div class="ss-ring" role="img" aria-label="Top ${topN} share ${_ssPct(share)}">
              <svg viewBox="0 0 128 128" aria-hidden="true">
                <circle cx="64" cy="64" r="${R}" class="trk"/>
                <circle cx="64" cy="64" r="${R}" class="arc" style="--c:${C.toFixed(1)};--off:${(C * (1 - share / 100)).toFixed(1)}"/>
              </svg>
              <div><b data-ss-count="${share.toFixed(1)}" data-ss-fmt="pct">${_ssPct(share)}</b><span>top ${topN}</span></div>
            </div>
            <div class="ss-stat"><span>${noun[0].toUpperCase() + noun.slice(1)}</span><b data-ss-count="${count}" data-ss-fmt="int">${_ssInt(count)}</b></div>
            ${qty ? `<div class="ss-stat"><span>Units</span><b data-ss-count="${Math.round(r.qty)}" data-ss-fmt="int">${_ssInt(r.qty)}</b></div>` : ''}
          </div>
        </div>
      </div>
    </section>

    <div class="ss-main">
      ${rows.length ? `
      <section class="ss-podium-wrap ss-reveal" id="ss-podium">
        <h2 class="ss-h">The podium</h2>
        <div class="ss-podium">${podOrder.map(n => { const x = pod[n]; return `
          <div class="ss-place p${n + 1}">
            <article class="ss-pod" role="button" tabindex="0" aria-label="Open ${escAttr(x.name)}" onclick="ssDetail(${n})" onkeydown="if(event.key==='Enter')ssDetail(${n})">
              ${n === 0 ? '<span class="ss-crown" aria-hidden="true">★</span>' : ''}
              <span class="ss-av lg">${escHtml(_ssInitials(x.name))}</span>
              <h3 class="ss-pod-nm">${escHtml(x.name)}</h3>
              ${items ? `<span class="px-code sm">${escHtml(x.key)}</span>` : ''}
              <b class="ss-pod-val" data-ss-count="${Math.round(x.a)}">${_ssGbp(x.a)}</b>
              <div class="ss-pod-meta">
                <span>${_ssPct(_ssShare(x.a, r.total))} of sales</span>
                ${cmp ? _ssChange(x.ty, x.ly) : ''}
              </div>
              ${qty ? `<div class="ss-pod-units">${_ssInt(x.q)} units</div>` : ''}
            </article>
            <div class="ss-step"><span>${n + 1}</span></div>
          </div>`; }).join('')}
        </div>
      </section>

      <section class="ss-mix ss-reveal">
        <div class="ss-mix-head">
          <h2 class="ss-h">${items ? 'Product split' : 'Customer split'}</h2>
          <p class="ss-sub">Each block is one of the top ${rows.length} ${noun}, sized by sales. The grey block is everyone else.</p>
        </div>
        <div class="ss-mixbar" role="img" aria-label="Top ${rows.length} ${_ssPct(share)}, other ${noun} ${_ssPct(100 - share)}">
          ${rows.map((x, n) => `<span class="ss-seg${n < 3 ? ' top' : ''}" style="--f:${_ssShare(x.a, r.total).toFixed(3)};--n:${n}" title="${escAttr(x.name)}: ${_ssGbp(x.a)} (${_ssPct(_ssShare(x.a, r.total))})"><em>${n + 1}</em></span>`).join('')}
          <span class="ss-seg rest" style="--f:${_ssShare(others, r.total).toFixed(3)};--n:${rows.length}" title="Other ${noun}: ${_ssGbp(others)}"><em>The other ${_ssInt(Math.max(0, count - rows.length))}</em></span>
        </div>
        <div class="ss-mix-legend">
          <span><b data-ss-count="${share.toFixed(1)}" data-ss-fmt="pct">${_ssPct(share)}</b> from the top ${rows.length}</span>
          <span><b data-ss-count="${Math.round(others)}">${_ssGbp(others)}</b> from the other ${_ssInt(Math.max(0, count - rows.length))} ${noun}</span>
        </div>
      </section>

      ${rest.length ? `
      <section class="ss-rest ss-reveal">
        <div class="ss-rank-head"><span>#</span><span>${items ? 'Product' : 'Customer'}</span><span>Share</span>${cmp ? `<span>${win} v ${yr - 1}</span>` : '<span></span>'}<span>Sales ${yr}</span></div>
        <ol class="ss-rank" start="4">${rest.map((x, n) => `
          <li style="--i:${n}" role="button" tabindex="0" onclick="ssDetail(${n + 3})" onkeydown="if(event.key==='Enter')ssDetail(${n + 3})">
            <span class="n">${String(n + 4).padStart(2, '0')}</span>
            <div class="who">
              <div class="nm">${escHtml(x.name)}${items ? ` <span class="px-code sm">${escHtml(x.key)}</span>` : ''}</div>
              <div class="ss-bar"><i style="--w:${(x.a / max * 100).toFixed(1)}%"></i></div>
            </div>
            <span class="sh">${_ssPct(_ssShare(x.a, r.total))}</span>
            <span class="ch">${cmp ? _ssChange(x.ty, x.ly) : ''}</span>
            <span class="v"><b data-ss-count="${Math.round(x.a)}">${_ssGbp(x.a)}</b>${qty ? `<small>${_ssInt(x.q)} units</small>` : ''}</span>
          </li>`).join('')}</ol>
      </section>` : ''}`
      : `<p class="ss-sub">No sales found this year for this topic. Check the Category or Items on the Spotlight list.</p>`}

      <section class="ss-panel ss-chart-panel ss-reveal">
        <div class="ss-chart-head">
          <div>
            <h3>The year so far</h3>
            <p>All customers, ${yr} against ${yr - 1}. ${SS_MONTHS[nm - 1]} is the month in progress${peak.m >= 0 ? `; the best complete month was ${SS_MONTHS[peak.m]} at ${_ssGbp(peak.v)}` : ''}. Hover a month for the figures.</p>
          </div>
          <div class="ss-chart-ctl">
            <div class="ss-toggle" role="group" aria-label="Chart view">
              <button data-mode="cum" aria-pressed="true" onclick="ssChartMode('cum')">Running total</button>
              <button data-mode="month" aria-pressed="false" onclick="ssChartMode('month')">Each month</button>
            </div>
            <div class="ss-key"><span><i class="ty"></i>${yr}</span><span><i class="ly"></i>${yr - 1}</span></div>
          </div>
        </div>
        <div id="ss-line-host" role="img" aria-label="${escAttr(w.sales)} by month, ${yr} against ${yr - 1}">${_ssLineChart(r.months, r.lyMonths, nm, yr, 'cum', r.closed < 12)}</div>
      </section>

      ${prev.length ? `
      <section class="ss-prev ss-reveal">
        <h2 class="ss-h">${isNow ? 'Previous spotlights' : 'Other spotlights'}</h2>
        <p class="ss-sub">Every past topic stays here. Opening one shows today's figures for it.</p>
        <div class="ss-grid">${prev.map(({ x, n }, k) => `
          <article class="px-card" style="--i:${k}" role="button" tabindex="0"
                   onclick="openSpotlight(${n})" onkeydown="if(event.key==='Enter')openSpotlight(${n})">
            <div class="px-card-media">
              ${_ssRange(x) ? `<span class="px-badge"><span class="px-badge-dot ${n === SS.cur ? 'green' : 'grey'}"></span>${escHtml(_ssRange(x))}</span>` : ''}
              <span class="px-card-initials">${escHtml(_ssTopicMark(x))}</span>
            </div>
            <div class="px-card-body">
              <div class="px-card-eyebrow">${n === SS.cur ? 'This fortnight' : escHtml(x.tag || x.category || 'Spotlight')}</div>
              <h3 class="px-card-title">${escHtml(x.title)}</h3>
            </div>
            <div class="px-card-go">Open</div>
          </article>`).join('')}</div>
      </section>` : ''}
    </div>`;
  _ssCountUp(host);
  _ssReveal(host);
  _ssStage();
}

// The beam: sweeps once on open, then follows the pointer over the
// dark stage. Pointer only — touch screens keep the sweep and no more.
function _ssStage() {
  const st = document.getElementById('ss-stage');
  if (!st) return;
  const still = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  if (still) return;
  let raf = 0;
  st.addEventListener('pointermove', e => {
    if (e.pointerType && e.pointerType !== 'mouse') return;
    const b = st.getBoundingClientRect();
    const x = ((e.clientX - b.left) / b.width * 100).toFixed(1), y = ((e.clientY - b.top) / b.height * 100).toFixed(1);
    cancelAnimationFrame(raf);
    raf = requestAnimationFrame(() => { st.style.setProperty('--mx', x + '%'); st.style.setProperty('--my', y + '%'); st.classList.add('ss-follow'); });
  });
  st.addEventListener('pointerleave', () => st.classList.remove('ss-follow'));
}

// ── The year so far: an animated line chart ───────────────────
// 28 Oct 2026, David didn't like the bars. Two lines, this year (ink,
// with a soft fill) and last year (grey), drawn left to right; dots pop
// in month by month; the month in progress is a hollow, pulsing dot.
// "Running total" shows the race; "Each month" shows the shape. Hover
// any month for the figures.
function _ssMono(pts) {
  // Monotone cubic through the points, so a line never overshoots a
  // month it passes through (Fritsch–Carlson).
  const n = pts.length;
  if (n < 2) return n ? `M${pts[0][0]},${pts[0][1]}` : '';
  const dx = [], dy = [], m = [];
  for (let i = 0; i < n - 1; i++) { dx[i] = pts[i + 1][0] - pts[i][0]; dy[i] = (pts[i + 1][1] - pts[i][1]) / dx[i]; }
  m[0] = dy[0]; m[n - 1] = dy[n - 2];
  for (let i = 1; i < n - 1; i++) m[i] = dy[i - 1] * dy[i] <= 0 ? 0 : (dy[i - 1] + dy[i]) / 2;
  for (let i = 0; i < n - 1; i++) {
    if (!dy[i]) { m[i] = m[i + 1] = 0; continue; }
    const a = m[i] / dy[i], b = m[i + 1] / dy[i], s = a * a + b * b;
    if (s > 9) { const t = 3 / Math.sqrt(s); m[i] = t * a * dy[i]; m[i + 1] = t * b * dy[i]; }
  }
  let d = `M${pts[0][0].toFixed(1)},${pts[0][1].toFixed(1)}`;
  for (let i = 0; i < n - 1; i++) {
    const h = dx[i] / 3;
    d += ` C${(pts[i][0] + h).toFixed(1)},${(pts[i][1] + m[i] * h).toFixed(1)} ${(pts[i + 1][0] - h).toFixed(1)},${(pts[i + 1][1] - m[i + 1] * h).toFixed(1)} ${pts[i + 1][0].toFixed(1)},${pts[i + 1][1].toFixed(1)}`;
  }
  return d;
}

function _ssK(n) {
  const a = Math.abs(n), sg = n < 0 ? '-' : '';
  if (a >= 1e6) return sg + '£' + (a / 1e6).toFixed(a >= 1e7 ? 0 : 1) + 'm';
  if (a >= 1e3) return sg + '£' + Math.round(a / 1e3) + 'k';
  return sg + '£' + Math.round(a);
}

function _ssLineChart(months, lyMonths, nm, yr, mode, partial) {
  const cum = a => a.reduce((o, x, k) => (o.push((o[k - 1] || 0) + x), o), []);
  const ty = (mode === 'cum' ? cum(months) : months.slice()).slice(0, nm);
  const ly = mode === 'cum' ? cum(lyMonths) : lyMonths.slice();
  const tyM = months.slice(0, nm), tyC = cum(months).slice(0, nm), lyC = cum(lyMonths);
  const W = 1000, H = 300, top = 16, bot = 12;
  const max = Math.max(1, ...ty, ...ly) * 1.1;
  const min = Math.min(0, ...ty, ...ly);
  const X = k => 20 + (W - 40) * (k / 11);
  const Y = v => top + (H - top - bot) * (1 - (v - min) / (max - min));
  const tyPts = ty.map((v, k) => [X(k), Y(v)]), lyPts = ly.map((v, k) => [X(k), Y(v)]);
  const base = Y(Math.max(0, min));
  const tyLine = _ssMono(tyPts), lyLine = _ssMono(lyPts);
  const area = tyPts.length > 1 ? `${tyLine} L${tyPts[tyPts.length - 1][0].toFixed(1)},${base.toFixed(1)} L${tyPts[0][0].toFixed(1)},${base.toFixed(1)} Z` : '';
  const grid = [0.25, 0.5, 0.75, 1].map(f => min + (max / 1.1 - min) * f);
  const last = ty.length - 1;
  const pct = (v, k) => `left:${(X(k) / W * 100).toFixed(2)}%;top:${(Y(v) / H * 100).toFixed(2)}%`;
  const dur = 1.5;

  return `
    <div class="ss-line ${mode}">
      <div class="ss-line-grid" aria-hidden="true">${grid.map(g => `
        <div style="top:${(Y(g) / H * 100).toFixed(2)}%"><span>${_ssK(g)}</span></div>`).join('')}
      </div>
      <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" aria-hidden="true">
        <defs>
          <linearGradient id="ss-fillg" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0" stop-color="#12100E" stop-opacity=".16"/>
            <stop offset="1" stop-color="#12100E" stop-opacity="0"/>
          </linearGradient>
        </defs>
        <g class="ly-g"><path class="ly" d="${lyLine}" vector-effect="non-scaling-stroke"/></g>
        <g class="ty-g" style="--dur:${dur}s">
          ${area ? `<path class="area" d="${area}"/>` : ''}
          <path class="ty" d="${tyLine}" vector-effect="non-scaling-stroke"/>
        </g>
      </svg>
      ${ty.map((v, k) => `<span class="ss-dot${k === last && partial ? ' now' : ''}" style="${pct(v, k)};--d:${(0.25 + dur * k / 11).toFixed(2)}s"></span>`).join('')}
      ${ly.map((v, k) => `<span class="ss-dot ly" style="${pct(v, k)};--d:${(0.1 + 0.9 * k / 11).toFixed(2)}s"></span>`).join('')}
      ${last >= 0 ? `<div class="ss-endlab" style="${pct(ty[last], last)};--d:${(0.3 + dur).toFixed(2)}s">
        <b>${_ssK(ty[last])}</b><span>${mode === 'cum' ? `${yr} so far` : `${SS_MONTHS[last]}${partial ? ' to date' : ''}`}</span></div>` : ''}
      <div class="ss-line-hit">${SS_MONTHS.map((mm, k) => `
        <div class="hc${k > last ? ' future' : ''}" style="left:${(X(k) / W * 100).toFixed(2)}%">
          <i class="guide"></i>
          <div class="tip ${k > 8 ? 'l' : k < 3 ? 'r' : ''}">
            <span class="mm">${mm}${k === last && partial ? ' · to date' : ''}</span>
            ${k <= last ? `<b>${_ssGbp(tyM[k])}</b><span>${yr} · running ${_ssK(tyC[k])}</span>` : `<b>–</b>`}
            <span class="lyv">${yr - 1}: ${_ssGbp(lyMonths[k])} · running ${_ssK(lyC[k])}</span>
          </div>
        </div>`).join('')}
      </div>
      <div class="ss-line-x" aria-hidden="true">${SS_MONTHS.map((mm, k) => `<span style="left:${(X(k) / W * 100).toFixed(2)}%">${mm}</span>`).join('')}</div>
    </div>`;
}

// Switch between running total and each month; the lines redraw.
function ssChartMode(mode) {
  const v = SS.view;
  const host = document.getElementById('ss-line-host');
  if (!v || !host) return;
  v.mode = mode;
  host.innerHTML = _ssLineChart(v.r.months, v.r.lyMonths, v.nm, v.yr, mode, v.r.closed < 12);
  document.querySelectorAll('.ss-toggle button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.mode === mode)));
}

// ── A customer (or product) up close ──────────────────────────
// Click anyone on the podium or in the ranks: a panel slides in with
// their months against last year, what they bought within the topic,
// and — for a customer — how big this topic is in their whole account
// with us. ‹ › (or the arrow keys) step through the ranking.
function _ssMatch(t) {
  const cats = String(t.category || '').split(/[;,]/).map(_ssSlug).filter(Boolean);
  const itemOk = _ssItemTest(t.items);
  return (cls, code) => (!cats.length || cats.indexOf(_ssSlug(cls)) >= 0) && (!itemOk || itemOk(code));
}

function _ssSum(a, from, to) { let s = 0; for (let k = from; k < to; k++) s += a[k] || 0; return s; }

function ssDetail(n) {
  const v = SS.view;
  if (!v || !v.r.rows[n]) return;
  const { t, r, data } = v;
  const x = r.rows[n];
  const ok = _ssMatch(t);
  const items = t.type === 'items';
  const yr = data.year, nm = _ssMonthCount(data), closed = r.closed;
  const qty = !!data.hasQty;

  const months = new Array(12).fill(0), lyMonths = new Array(12).fill(0);
  const lines = new Map();               // what makes up their total

  // Only the spotlighted products: David, 28 Oct — "when you open the
  // customer it's just the spotlighted, for example just LFX".
  for (const [comp, classes] of Object.entries(data.c || {})) {
    if (!items && comp !== x.key) continue;
    for (const [cls, its] of Object.entries(classes)) {
      for (const [code, val] of Object.entries(its)) {
        const tyA = _ssSum(val.a, 12, 24);
        const hit = items ? code === x.key && ok(cls, code) : ok(cls, code);
        if (!hit) continue;
        for (let k = 0; k < 12; k++) { months[k] += val.a[12 + k] || 0; lyMonths[k] += val.a[k] || 0; }
        const key = items ? comp : code;
        const e = lines.get(key) || { key, a: 0, q: 0, name: items ? comp : ((data.names || {})[code] || code) };
        e.a += tyA; e.q += val.q ? _ssSum(val.q, 12, 24) : 0;
        lines.set(key, e);
      }
    }
  }
  const list = [...lines.values()].filter(e => e.a > 0).sort((p, q) => q.a - p.a);
  const lmax = list.length ? list[0].a : 1;
  const w = _ssWords(t);
  const cMax = Math.max(1, ...months.slice(0, nm), ...lyMonths);
  const total = r.rows.length;

  let el = document.getElementById('ss-drawer');
  if (!el) {
    el = document.createElement('div');
    el.id = 'ss-drawer';
    el.setAttribute('role', 'dialog');
    el.setAttribute('aria-modal', 'true');
  }
  const fresh = !el.classList.contains('on');
  const step = fresh ? 0 : (n > SS.detail ? 1 : -1);
  // Fullscreen only draws the Present element and what is inside it.
  const pres = document.documentElement.classList.contains('ss-presenting') && document.getElementById('ss-present');
  const home = pres || document.body;
  if (el.parentNode !== home) home.appendChild(el);
  el.inert = false; el.removeAttribute('data-ss-inert');
  el.setAttribute('aria-label', x.name);
  el.innerHTML = `
    <div class="ss-dr-shade" onclick="ssDetailClose()"></div>
    <aside class="ss-dr">
      <header class="ss-dr-head">
        <div class="ss-dr-nav">
          <button onclick="ssDetail(${n - 1})" ${n === 0 ? 'disabled' : ''} aria-label="Previous">‹</button>
          <span>#${n + 1} of ${_ssInt(total)}</span>
          <button onclick="ssDetail(${n + 1})" ${n >= total - 1 ? 'disabled' : ''} aria-label="Next">›</button>
        </div>
        <button class="ss-dr-x" onclick="ssDetailClose()" aria-label="Close">×</button>
      </header>
      <div class="ss-dr-body">
        <div class="ss-dr-who">
          <span class="ss-av xl">${escHtml(_ssInitials(x.name))}</span>
          <div>
            <div class="px-eyebrow">${escHtml(w.sales)} · ${yr} to date</div>
            <h2>${escHtml(x.name)}</h2>
            ${items ? `<span class="px-code sm">${escHtml(x.key)}</span>` : ''}
          </div>
        </div>
        <div class="ss-dr-fig">
          <b class="ss-big" data-ss-count="${Math.round(x.a)}">${_ssGbp(x.a)}</b>
          <div class="ss-dr-chips">
            ${closed ? `${_ssChange(x.ty, x.ly, { long: true })}<span class="ss-dr-win">${_ssWindow(closed)} v ${yr - 1}</span>` : ''}
          </div>
          <div class="ss-dr-stats">
            <div><span>Share of ${escHtml(w.salesLc)}</span><b>${_ssPct(_ssShare(x.a, r.total))}</b></div>
            ${qty ? `<div><span>Units</span><b data-ss-count="${Math.round(x.q)}" data-ss-fmt="int">${_ssInt(x.q)}</b></div>` : ''}
            <div><span>${yr - 1} full year</span><b>${_ssGbp(x.lyAll)}</b></div>
          </div>
        </div>

        <section class="ss-dr-sec">
          <h3>Month by month</h3>
          <div class="ss-dr-chart" role="img" aria-label="${escAttr(x.name)} by month">
            ${SS_MONTHS.map((mm, k) => `
            <div class="c${k === nm - 1 ? ' part' : ''}" title="${mm}: ${k < nm ? _ssGbp(months[k]) + ' · ' : ''}${yr - 1}: ${_ssGbp(lyMonths[k])}">
              <div class="p"><i class="ly" style="--h:${(Math.max(0, lyMonths[k]) / cMax * 100).toFixed(1)}%;--k:${k}"></i><i class="ty" style="--h:${k < nm ? (Math.max(0, months[k]) / cMax * 100).toFixed(1) : 0}%;--k:${k}"></i></div>
              <span>${mm[0]}</span>
            </div>`).join('')}
          </div>
          <div class="ss-key sm"><span><i class="ty"></i>${yr}</span><span><i class="ly"></i>${yr - 1}</span></div>
        </section>

        <section class="ss-dr-sec">
          <h3>${items ? 'Who buys it' : 'What they bought'}</h3>
          <ol class="ss-dr-list">${list.slice(0, 12).map((e, k) => `
            <li style="--k:${k}">
              <div class="nm"><b>${escHtml(e.name)}</b>${items ? '' : `<span>${escHtml(e.key)}</span>`}</div>
              <div class="v">${_ssGbp(e.a)}${qty ? `<small>${_ssInt(e.q)} units</small>` : ''}</div>
              <div class="ss-bar"><i style="--w:${(e.a / lmax * 100).toFixed(1)}%;--i:${k}"></i></div>
            </li>`).join('')}</ol>
          ${list.length > 12 ? `<p class="ss-dr-more">and ${list.length - 12} more</p>` : ''}
        </section>
      </div>
    </aside>`;
  if (fresh) {
    SS.drawerFocus = document.activeElement;
    document.addEventListener('keydown', _ssDetailKey, true);   // capture: Esc closes this, not present mode under it
    Array.from(home.children).forEach(nd => { if (nd !== el && !nd.inert) { nd.inert = true; nd.setAttribute('data-ss-dr-inert', ''); } });
    if (pres) Array.from(document.body.children).forEach(nd => { if (nd !== pres && !nd.inert) { nd.inert = true; nd.setAttribute('data-ss-dr-inert', ''); } });
    document.documentElement.classList.add('ss-drawer-open');
  }
  el.classList.add('on');
  SS.detail = n;
  _ssCountUp(el);
  // Stepping redraws the panel, so put focus back on the arrow used
  // (or on Close at either end of the ranking).
  const nav = el.querySelectorAll('.ss-dr-nav button');
  const want = step > 0 ? nav[1] : step < 0 ? nav[0] : null;
  (want && !want.disabled ? want : el.querySelector('.ss-dr-x')).focus();
}

function ssDetailClose() {
  const el = document.getElementById('ss-drawer');
  if (!el || !el.classList.contains('on')) return;
  el.classList.remove('on');
  document.removeEventListener('keydown', _ssDetailKey, true);
  document.querySelectorAll('[data-ss-dr-inert]').forEach(nd => { nd.inert = false; nd.removeAttribute('data-ss-dr-inert'); });
  document.documentElement.classList.remove('ss-drawer-open');
  const back = SS.drawerFocus; SS.drawerFocus = null;
  if (back && back.focus && document.contains(back)) back.focus();
}

function _ssDetailKey(e) {
  if (e.key === 'Escape') { e.stopImmediatePropagation(); ssDetailClose(); }
  else if (e.key === 'ArrowRight') ssDetail(SS.detail + 1);
  else if (e.key === 'ArrowLeft') ssDetail(SS.detail - 1);
}

// Sections below the fold hold their entrance until they scroll into
// view. Without IntersectionObserver they simply play straight away.
function _ssReveal(root) {
  const els = (root || document).querySelectorAll('.ss-reveal');
  if (!('IntersectionObserver' in window)) { els.forEach(e => { e.classList.add('ss-seen'); _ssCountUp(e); }); return; }
  const io = new IntersectionObserver(ents => ents.forEach(en => {
    if (en.isIntersecting) { en.target.classList.add('ss-seen'); io.unobserve(en.target); _ssCountUp(en.target); }
  }), { threshold: 0, rootMargin: '0px 0px -10% 0px' });
  els.forEach(e => io.observe(e));
  // Never leave a section waiting: a tall one, a background tab or an
  // observer that never fires still shows within two seconds.
  setTimeout(() => els.forEach(e => { if (!e.classList.contains('ss-seen')) { e.classList.add('ss-seen'); _ssCountUp(e); } }), 2000);
}

// ── Present ───────────────────────────────────────────────────
// For the Monday Sales & Marketing meeting: the spotlight on its own,
// full screen, big enough to read across a room. Esc or the button
// closes it. Built from the same figures as the page.
function ssPresent() {
  const t = SS.topics[SS.open];
  if (!t || !SS.data) return;
  const data = SS.data, r = _ssCompute(t, data), w = _ssWords(t);
  const topN = (HUB_CONFIG.spotlight || {}).topN || 10;
  const rows = r.rows.slice(0, topN), max = rows.length ? rows[0].a : 1;
  const cmp  = r.closed > 0, win = _ssWindow(r.closed);

  let el = document.getElementById('ss-present');
  if (!el) {
    el = document.createElement('div');
    el.id = 'ss-present';
    el.setAttribute('role', 'dialog');
    el.setAttribute('aria-modal', 'true');
    document.body.appendChild(el);
  }
  el.setAttribute('aria-label', t.title);
  el.innerHTML = `
    <button class="ss-pr-close" onclick="ssPresentClose()" aria-label="Close">Close <kbd>Esc</kbd></button>
    <div class="ss-pr-head">
      <div>
        <div class="px-eyebrow">Sales Spotlight</div>
        <h1>${escHtml(t.title)}</h1>
      </div>
      <div class="ss-pr-fig">
        <b class="ss-big" data-ss-count="${Math.round(r.total)}">${_ssGbp(r.total)}</b>
        <span>${escHtml(w.sales)} · ${data.year} to date${cmp && r.likeLy ? ` · ${win} v ${data.year - 1} ` : ''}${cmp && r.likeLy ? _ssChange(r.likeTy, r.likeLy) : ''}</span>
      </div>
    </div>
    <ol class="ss-pr-list">${rows.map((x, n) => `
      <li style="--i:${n}" role="button" tabindex="0" onclick="ssDetail(${n})" onkeydown="if(event.key==='Enter')ssDetail(${n})">
        <span class="n">${n + 1}</span>
        <span class="nm">${escHtml(x.name)}</span>
        <span class="ss-bar"><i style="--w:${(x.a / max * 100).toFixed(1)}%"></i></span>
        <span class="v">${_ssGbp(x.a)}</span>
        <span class="c">${cmp ? _ssChange(x.ty, x.ly) : ''}</span>
      </li>`).join('')}</ol>
    <div class="ss-pr-foot">NetSuite, via the sales dashboard${SS.file && SS.file.modified ? ` · updated ${escHtml(fmtSpDate(SS.file.modified))}` : ''}${cmp ? ` · change is ${win} ${data.year} v ${win} ${data.year - 1}` : ''}</div>`;
  SS.returnFocus = document.activeElement;
  el.classList.add('on');
  document.documentElement.classList.add('ss-presenting');
  document.addEventListener('keydown', _ssPresentKey);
  // The page behind can't be reached with Tab while presenting.
  document.querySelectorAll('body > *:not(#ss-present):not(#ss-drawer)').forEach(n => { if (!n.inert) { n.inert = true; n.setAttribute('data-ss-inert', ''); } });
  try { const p = el.requestFullscreen && el.requestFullscreen(); if (p && p.catch) p.catch(() => {}); } catch (e) { /* stays a full-window view */ }
  _ssCountUp(el);
  const btn = el.querySelector('.ss-pr-close'); if (btn) btn.focus();
}

function ssPresentClose() {
  ssDetailClose();
  const el = document.getElementById('ss-present');
  if (el) el.classList.remove('on');
  document.documentElement.classList.remove('ss-presenting');
  document.removeEventListener('keydown', _ssPresentKey);
  document.querySelectorAll('[data-ss-inert]').forEach(n => { n.inert = false; n.removeAttribute('data-ss-inert'); });
  try { const p = document.fullscreenElement && document.exitFullscreen && document.exitFullscreen(); if (p && p.catch) p.catch(() => {}); } catch (e) { /* already out */ }
  const back = SS.returnFocus; SS.returnFocus = null;
  if (back && back.focus && document.contains(back)) back.focus();
}

function _ssPresentKey(e) { if (e.key === 'Escape') ssPresentClose(); }

// Leaving full screen with the browser's own Esc closes the view too.
document.addEventListener('fullscreenchange', () => {
  if (!document.fullscreenElement && document.documentElement.classList.contains('ss-presenting')) { ssDetailClose(); ssPresentClose(); }
});
