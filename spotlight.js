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
    const item   = await graphFetch(`/sites/${siteId}/drive/root:/${path}?$select=id,lastModifiedDateTime,@microsoft.graph.downloadUrl`);
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
function _ssCompute(t, data) {
  const cats    = String(t.category || '').split(/[;,]/).map(_ssSlug).filter(Boolean);
  const itemOk  = _ssItemTest(t.items);
  const out     = { rows: [], months: new Array(12).fill(0), total: 0, qty: 0, lastYear: 0, buyers: 0 };
  if (!cats.length && !itemOk) return out;

  const by = new Map();
  const buyers = new Set();
  for (const [comp, classes] of Object.entries(data.c || {})) {
    for (const [cls, items] of Object.entries(classes)) {
      if (cats.length && cats.indexOf(_ssSlug(cls)) < 0) continue;
      for (const [code, v] of Object.entries(items)) {
        if (itemOk && !itemOk(code)) continue;
        let a = 0, q = 0;
        for (let m = 0; m < 12; m++) {
          a += v.a[12 + m] || 0;
          q += (v.q && v.q[12 + m]) || 0;
          out.months[m] += v.a[12 + m] || 0;
          out.lastYear  += v.a[m] || 0;
        }
        if (!a && !q) continue;
        const key = t.type === 'items' ? code : comp;
        const e = by.get(key) || { key, a: 0, q: 0 };
        e.a += a; e.q += q;
        by.set(key, e);
        out.total += a; out.qty += q;
        if (a > 0) buyers.add(comp);
      }
    }
  }
  out.buyers = buyers.size;
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

  const r = _ssCompute(t, data);
  const top3 = r.rows.slice(0, 3);
  const w = _ssWords(t);
  const range = _ssRange(t);

  host.innerHTML = `
   <div class="pp-band ss-band">
    <div class="pp-band-head">
      <h2 class="pp-band-title">Sales Spotlight</h2>
      <span class="pp-band-note">A new topic every fortnight</span>
    </div>
    <article class="px-card ss-feature" role="button" tabindex="0"
             onclick="openSpotlight(${SS.cur})" onkeydown="if(event.key==='Enter')openSpotlight(${SS.cur})">
      <div class="px-card-media">
        ${range ? `<span class="px-badge"><span class="px-badge-dot green"></span>${escHtml(range)}</span>` : ''}
        <div class="ss-fig"><b>${_ssGbp(r.total)}</b><span>${escHtml(w.sales)} this year</span></div>
      </div>
      <div class="px-card-body">
        <div class="px-card-eyebrow">This fortnight${t.category ? ' · ' + escHtml(t.category) : ''}</div>
        <h3 class="px-card-title">${escHtml(t.title)}</h3>
        ${top3.length ? `<ol class="ss-top3">${top3.map((x, i) => `
          <li><b>${i + 1}</b><span>${escHtml(x.name)}</span><i>${_ssGbp(x.a)}</i></li>`).join('')}</ol>`
          : `<p class="px-card-meta">No sales found this year. Check the Category or Items on the Spotlight list.</p>`}
        <div class="px-card-go">See the full top ${(HUB_CONFIG.spotlight || {}).topN || 10}
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="13" height="13"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
        </div>
      </div>
    </article>
   </div>`;
  if (!r.rows.length) console.warn(`[Spotlight] "${t.title}" matched no sales — Category "${t.category}", Items "${t.items}".`);
}

// ── The Spotlight page ────────────────────────────────────────
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

  const topN  = (HUB_CONFIG.spotlight || {}).topN || 10;
  const r     = _ssCompute(t, data);
  const rows  = r.rows.slice(0, topN);
  const max   = rows.length ? rows[0].a : 1;
  const nm    = _ssMonthCount(data);
  const mVals = r.months.slice(0, nm);
  const mMax  = Math.max(1, ...mVals);
  const w     = _ssWords(t);
  const yr    = data.year;
  const range = _ssRange(t);
  const isNow = i === SS.cur;
  const upd   = SS.file && SS.file.modified ? fmtSpDate(SS.file.modified) : '';
  const noun  = t.type === 'items' ? 'products' : 'customers';
  const count = t.type === 'items' ? r.rows.length : r.buyers;

  const prev = SS.topics.map((x, n) => ({ x, n })).filter(o => o.n !== i);

  host.innerHTML = `
    <section class="px-lead dark">
      <div class="px-lead-copy">
        <div class="px-eyebrow">Sales Spotlight</div>
        <h1 class="px-lead-title">${escHtml(t.title)}</h1>
        <p class="px-lead-sub">Our top ${topN} ${noun} by ${escHtml(w.salesLc)}, 1 January to date. Refreshed from NetSuite each weekday morning by the sales dashboard; the date is under About this spotlight.</p>
        <div class="px-lead-codes">
          <span class="px-code">${_ssInt(count)} ${noun}</span>
          ${data.hasQty ? `<span class="px-code">${_ssInt(r.qty)} units</span>` : ''}
          <span class="px-code">Full year ${yr - 1}: ${_ssGbp(r.lastYear)}</span>
        </div>
        ${range ? `<div class="px-lead-meta">
          <span class="px-badge"><span class="px-badge-dot ${isNow ? 'green' : 'grey'}"></span>${isNow ? 'This fortnight' : 'Previous spotlight'}</span>
          <span class="px-lead-when">${escHtml(range)}</span>
        </div>` : ''}
      </div>
      <div class="px-lead-media">
        <div class="ss-fig lg"><b>${_ssGbp(r.total)}</b><span>${escHtml(w.sales)} · ${yr} to date</span></div>
      </div>
    </section>

    <div class="ss-main">
      <div class="ss-cols">
        <div>
          <h2 class="ss-h">Top ${topN} ${escHtml(w.adj)}${noun}</h2>
          <p class="ss-sub">Ranked by sales so far this year${isNow ? '' : '. Today\'s figures, not the ones shown at the time'}.</p>
          ${rows.length ? `
          <div class="ss-rank-head"><span>#</span><span>${t.type === 'items' ? 'Product' : 'Customer'}</span><span>Sales ${yr}</span></div>
          <ol class="ss-rank">${rows.map((x, n) => `
            <li>
              <span class="n">${String(n + 1).padStart(2, '0')}</span>
              <div>
                <div class="nm">${escHtml(x.name)}${t.type === 'items' ? ` <span class="px-code sm">${escHtml(x.key)}</span>` : ''}</div>
                <div class="bar"><i style="width:${(x.a / max * 100).toFixed(1)}%"></i></div>
              </div>
              <span class="v">${_ssGbp(x.a)}${data.hasQty ? `<small>${_ssInt(x.q)} units</small>` : ''}</span>
            </li>`).join('')}</ol>`
          : `<p class="ss-sub">No sales found this year for this topic. Check the Category or Items on the Spotlight list.</p>`}
        </div>

        <aside class="ss-aside">
          <div class="ss-panel">
            <h3>${escHtml(w.chart)}</h3>
            <p>All customers, ${yr}.</p>
            <div class="ss-chart" role="img" aria-label="${escAttr(w.sales)} by month, ${yr}">
              ${mVals.map((v, m) => `<div class="${m === nm - 1 ? 'part' : ''}" title="${SS_MONTHS[m]}: ${_ssGbp(v)}"><i style="height:${Math.max(0, v / mMax * 100).toFixed(1)}%"></i></div>`).join('')}
            </div>
            <div class="ss-chart-lab">${mVals.map((_, m) => `<span>${SS_MONTHS[m].charAt(0)}</span>`).join('')}</div>
            <div class="ss-note">${SS_MONTHS[nm - 1]} is the month in progress.</div>
          </div>
          <div class="ss-panel">
            <h3>About this spotlight</h3>
            <dl>
              ${range ? `<dt>Runs</dt><dd>${escHtml(range)}</dd>` : ''}
              ${upd ? `<dt>Updated</dt><dd>${escHtml(upd)}</dd>` : ''}
              <dt>Source</dt><dd>NetSuite, via the sales dashboard</dd>
              <dt>Measure</dt><dd>Sales £, net of credits</dd>
            </dl>
          </div>
        </aside>
      </div>

      ${prev.length ? `
      <div class="ss-prev">
        <h2 class="ss-h">${isNow ? 'Previous spotlights' : 'Other spotlights'}</h2>
        <p class="ss-sub">Every past topic stays here. Opening one shows today's figures for it.</p>
        <div class="ss-grid">${prev.map(({ x, n }, k) => `
          <article class="px-card" style="--i:${k}" role="button" tabindex="0"
                   onclick="openSpotlight(${n})" onkeydown="if(event.key==='Enter')openSpotlight(${n})">
            <div class="px-card-media">
              ${_ssRange(x) ? `<span class="px-badge"><span class="px-badge-dot ${n === SS.cur ? 'green' : 'grey'}"></span>${escHtml(_ssRange(x))}</span>` : ''}
              <span class="px-card-initials">${escHtml((x.category || x.title).slice(0, 3).toUpperCase())}</span>
            </div>
            <div class="px-card-body">
              <div class="px-card-eyebrow">${n === SS.cur ? 'This fortnight' : escHtml(x.category || 'Spotlight')}</div>
              <h3 class="px-card-title">${escHtml(x.title)}</h3>
            </div>
            <div class="px-card-go">Open</div>
          </article>`).join('')}</div>
      </div>` : ''}
    </div>`;
}
