/**
 * CheckFire Marketing Hub — Data layer
 * ─────────────────────────────────────────────────────────────
 * 1. WordPress public feed (no auth) — cached so the hero grid
 *    and news section share a single network request.
 * 2. Microsoft Graph → SharePoint lists & document library.
 *    Requires getAccessToken() from auth.js.
 *
 * All dynamic values are escaped via escHtml/escAttr/safeUrl
 * (defined in ui.js) before touching innerHTML.
 */

// ═══ WordPress News ══════════════════════════════════════════

let _wpPromise = null;

function fetchWordPressNews() {
  if (_wpPromise) return _wpPromise;

  _wpPromise = (async () => {
    const { apiUrl, postsPerPage } = HUB_CONFIG.wordpress;
    const url = `${apiUrl}/posts?per_page=${postsPerPage}&_fields=id,title,excerpt,date,link,jetpack_featured_media_url,_links&_embed=wp:featuredmedia`;

    const res = await fetch(url);
    if (!res.ok) throw new Error(`WordPress API returned ${res.status}`);
    const posts = await res.json();

    return posts.map(post => {
      let image = post.jetpack_featured_media_url || null;
      if (!image && post._embedded?.['wp:featuredmedia']?.[0]?.source_url) {
        image = post._embedded['wp:featuredmedia'][0].source_url;
      }

      const excerpt = (post.excerpt?.rendered || '')
        .replace(/<[^>]+>/g, '')
        .replace(/\[&hellip;\]/g, '…')
        .replace(/&#8217;/g, "'")
        .trim()
        .slice(0, 160);

      return {
        id:      post.id,
        title:   (post.title?.rendered || 'Untitled').replace(/&#8217;/g, "'").replace(/&amp;/g, '&'),
        excerpt,
        date:    post.date ? new Date(post.date).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : '',
        link:    post.link || '#',
        image,
      };
    });
  })();

  // A failed fetch shouldn't poison the cache — allow retry
  _wpPromise.catch(() => { _wpPromise = null; });

  return _wpPromise;
}

// WordPress "pages" → the home-page "Updated Landing Pages" carousel.
// Public endpoint, no auth. Sorted newest-modified first.
let _wpPagesPromise = null;

function fetchWordPressPages() {
  if (_wpPagesPromise) return _wpPagesPromise;

  _wpPagesPromise = (async () => {
    const { apiUrl } = HUB_CONFIG.wordpress;
    const per    = (HUB_CONFIG.wordpress.pagesPerPage) || 8;
    const parent = (HUB_CONFIG.wordpress.landingPageParent) || 0;
    let url = `${apiUrl}/pages?per_page=${per}&orderby=modified&order=desc&_fields=id,title,excerpt,modified,link,jetpack_featured_media_url,_links&_embed=wp:featuredmedia`;
    if (parent) url += `&parent=${encodeURIComponent(parent)}`;

    const res = await fetch(url);
    if (!res.ok) throw new Error(`WordPress pages API returned ${res.status}`);
    const pages = await res.json();

    return (pages || []).map(p => {
      let image = p.jetpack_featured_media_url || null;
      if (!image && p._embedded?.['wp:featuredmedia']?.[0]?.source_url) {
        image = p._embedded['wp:featuredmedia'][0].source_url;
      }
      const excerpt = (p.excerpt?.rendered || '')
        .replace(/<[^>]+>/g, '').replace(/\[&hellip;\]/g, '…')
        .replace(/&#8217;/g, "'").trim().slice(0, 140);
      return {
        id:      p.id,
        title:   (p.title?.rendered || 'Untitled').replace(/&#8217;/g, "'").replace(/&amp;/g, '&'),
        excerpt,
        date:    p.modified ? new Date(p.modified).toLocaleDateString('en-GB', { day:'numeric', month:'short', year:'numeric' }) : '',
        link:    p.link || '#',
        image,
      };
    });
  })();

  _wpPagesPromise.catch(() => { _wpPagesPromise = null; });
  return _wpPagesPromise;
}

// ═══ Microsoft Graph — shared plumbing ═══════════════════════

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_CACHE_TTL = 5 * 60 * 1000; // 5 minutes

function _cacheGet(key) {
  try {
    const raw = sessionStorage.getItem('hubcache_' + key);
    if (!raw) return null;
    const { t, v } = JSON.parse(raw);
    return (Date.now() - t < GRAPH_CACHE_TTL) ? v : null;
  } catch (_) { return null; }
}

function _cacheSet(key, v) {
  try { sessionStorage.setItem('hubcache_' + key, JSON.stringify({ t: Date.now(), v })); }
  catch (_) { /* storage full / private mode — fine, just uncached */ }
}

async function graphFetch(path) {
  const token = await getAccessToken();
  if (!token) throw new Error('Not signed in');

  const res = await fetch(GRAPH_BASE + path, {
    headers: { Authorization: 'Bearer ' + token },
  });

  if (res.status === 404) throw new Error('NOT_FOUND');
  if (res.status === 403) throw new Error('Permission denied — has admin consent been granted?');
  if (!res.ok) throw new Error('Graph returned ' + res.status);
  return res.json();
}

// POST variant — used for the document "preview" action, which returns a
// short-lived, embeddable URL so files open INSIDE the hub (not SharePoint).
async function graphPost(path, body) {
  const token = await getAccessToken();
  if (!token) throw new Error('Not signed in');
  const res = await fetch(GRAPH_BASE + path, {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify(body || {}),
  });
  if (!res.ok) throw new Error('Graph POST returned ' + res.status);
  return res.json();
}

// Resolve a SharePoint site ID from its URL, once per site per session.
// The hub now talks to more than one site (MarketingHub + Product
// Portal), so resolution is keyed by URL rather than a single global.
const _siteIdPromises = {};

function resolveSiteId(siteUrl) {
  const key = siteUrl || HUB_CONFIG.sharepointSite;
  if (_siteIdPromises[key]) return _siteIdPromises[key];

  _siteIdPromises[key] = (async () => {
    const ck = 'siteId_' + key;
    const cached = _cacheGet(ck);
    if (cached) return cached;

    const u = new URL(key);
    const data = await graphFetch(`/sites/${u.hostname}:${u.pathname}`);
    _cacheSet(ck, data.id);
    return data.id;
  })();

  _siteIdPromises[key].catch(() => { delete _siteIdPromises[key]; });
  return _siteIdPromises[key];
}

// Back-compat: the default (MarketingHub) site used by the list fetchers.
function getSiteId() {
  return resolveSiteId(HUB_CONFIG.sharepointSite);
}

// Resolve a document library ("drive") on a given site by name, falling
// back to the site's first drive. Used by the in-hub file browser for
// both the Marketing library and the Product Portal.
async function resolveDrive(siteUrl, libraryName) {
  const siteId = await resolveSiteId(siteUrl);
  const drives = await graphFetch(`/sites/${siteId}/drives?$select=id,name`);
  const wanted = (libraryName || 'Documents').toLowerCase();
  const drive  = (drives.value || []).find(d => (d.name || '').toLowerCase() === wanted)
              || (drives.value || [])[0];
  if (!drive) throw new Error(`No document library found on ${siteUrl}`);
  return drive;
}

// Children of a drive folder (root when itemId is null). Each item is
// stamped with its drive id so previews can build the /preview path.
async function fetchDriveChildren(driveId, itemId) {
  const base = itemId
    ? `/drives/${driveId}/items/${itemId}/children`
    : `/drives/${driveId}/root/children`;
  const data = await graphFetch(
    `${base}?$select=id,name,size,lastModifiedDateTime,webUrl,file,folder&$top=200`
  );
  return (data.value || []).map(f => ({ ...f, _driveId: driveId }));
}

// ═══ Fetchers ════════════════════════════════════════════════

async function fetchListItems(listName) {
  const cacheKey = 'list_' + listName;
  const cached = _cacheGet(cacheKey);
  if (cached) return cached;

  const siteId = await getSiteId();
  try {
    const data = await graphFetch(
      `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items?expand=fields&$top=100`
    );
    const items = (data.value || []).map(i => i.fields || {});
    _cacheSet(cacheKey, items);
    return items;
  } catch (e) {
    if (e.message === 'NOT_FOUND') {
      throw new Error(`List "${listName}" not found — check the name in config.js (case-sensitive)`);
    }
    throw e;
  }
}

async function fetchLibraryFiles() {
  const cached = _cacheGet('library');
  if (cached) return cached;

  const siteId = await getSiteId();
  const drives = await graphFetch(`/sites/${siteId}/drives?$select=id,name`);
  const wanted = (HUB_CONFIG.documentsLibrary || 'Documents').toLowerCase();
  const drive  = (drives.value || []).find(d => (d.name || '').toLowerCase() === wanted)
              || (drives.value || [])[0];
  if (!drive) throw new Error(`Document library "${HUB_CONFIG.documentsLibrary}" not found`);

  const data = await graphFetch(
    `/drives/${drive.id}/root/children?$select=id,name,size,lastModifiedDateTime,webUrl,file,folder&$top=100`
  );
  // Stash the drive id on each item so the in-hub preview can build the
  // /drives/{drive}/items/{item}/preview path later.
  const files = (data.value || []).map(f => ({ ...f, _driveId: drive.id }));
  _cacheSet('library', files);
  return files;
}

// ═══ Formatting helpers ══════════════════════════════════════

function fmtSpDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return '';
  return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
}

function fmtMoney(v) {
  const n = Number(v);
  if (isNaN(n)) return escHtml(v);
  return '£' + n.toLocaleString('en-GB');
}

function humanSize(bytes) {
  const n = Number(bytes);
  if (!n) return '';
  if (n < 1024) return n + ' B';
  if (n < 1048576) return (n / 1024).toFixed(0) + ' KB';
  return (n / 1048576).toFixed(1) + ' MB';
}

// SharePoint hyperlink columns arrive as { Url, Description }
function linkOf(v) {
  return (v && typeof v === 'object') ? v.Url : v;
}

function statusBadge(status) {
  const s = String(status || '').toLowerCase();
  let tone = '';
  if (/live|complete|confirmed|available|launched/.test(s))      tone = 'green';
  else if (/planning|review|upcoming|draft/.test(s))             tone = 'amber';
  else if (/pending|delayed|cancelled|blocked/.test(s))          tone = 'red';
  if (!status) return '';
  return `<span class="badge ${tone}">${tone ? `<span class="status-dot ${tone}"></span>` : ''}${escHtml(status)}</span>`;
}

function fileIcon(name, isFolder) {
  if (isFolder) return { cls: 'doc', label: 'DIR' };
  const ext = String(name).split('.').pop().toLowerCase();
  if (['png','jpg','jpeg','gif','svg','webp'].includes(ext)) return { cls: 'img', label: 'IMG' };
  if (['mp4','mov','avi','webm'].includes(ext))              return { cls: 'vid', label: ext.toUpperCase() };
  return { cls: 'doc', label: ext.slice(0, 4).toUpperCase() || 'FILE' };
}

// ═══ RAG status (traffic lights) ═════════════════════════════
// One shared lifecycle for product launches AND campaigns, so the
// colour means the same thing wherever you see it:
//   red   — planning / drafting / on hold / delayed
//   amber — scheduled or upcoming (work in flight)
//   green — launched / live
//   grey  — finished, archived, or a status we don't recognise
function ragOf(status) {
  const t = String(status || '').toLowerCase();
  if (!t) return 'grey';
  if (/complete|closed|archiv|finished|ended/.test(t))                 return 'grey';
  if (/launch|live|released|active|published/.test(t))                 return 'green';
  if (/schedul|upcoming|confirm|ready|approved|in progress/.test(t))   return 'amber';
  if (/plan|draft|concept|hold|delay|pending|plan/.test(t))            return 'red';
  return 'grey';
}

function ragChip(status) {
  const tone = ragOf(status);
  return `<span class="rag ${tone}"><span class="rag-dot"></span>${escHtml(status || 'Not set')}</span>`;
}

// Badge that sits on a card thumbnail (absolute-positioned .pill)
function ragPill(status) {
  const tone = ragOf(status);
  return `<span class="pill"><span class="status-dot ${tone}"></span>${escHtml(status || 'Not set')}</span>`;
}

// Product codes: SKU columns often hold several codes separated by
// commas, semicolons or slashes. Split so each renders as its own chip.
function productCodes(f) {
  const raw = [f.SKU, f.ProductCode, f.ProductCodes, f.Codes].find(v => v !== undefined && v !== null && v !== '');
  if (!raw) return [];
  return String(raw).split(/[,;/\n]+/).map(x => x.trim()).filter(Boolean);
}

// ═══ Renderers ═══════════════════════════════════════════════

let _launchItems = [];

// Product Launches page. Follows the same card layout as Campaigns
// (marketing asked for the two pages to match) and colours each card by
// its RAG stage. The header panel counts launches by stage.
function renderLaunches(items) {
  const el = document.getElementById('sp-launches-list');
  if (!el) return;

  items = items || [];

  // Header panel — counted from live data, zeros when the list is empty.
  const byTone = t => items.filter(f => ragOf(f.Status) === t).length;
  const _set = (id, v) => { const n = document.getElementById(id); if (n) n.textContent = v; };
  _set('sp-metric-launch-planning',  byTone('red'));
  _set('sp-metric-launch-scheduled', byTone('amber'));
  _set('sp-metric-launch-live',      byTone('green'));
  _set('sp-metric-launch-total',     items.length);

  if (!items.length) {
    el.innerHTML = '<p class="prose dim">No launches in SharePoint yet — add items to the Product Launches list.</p>';
    return;
  }

  // Newest first so the current launch leads the page.
  const sorted = [...items].sort((a, b) => String(b.LaunchDate || '').localeCompare(String(a.LaunchDate || '')));
  _launchItems = sorted;

  el.innerHTML = `<div class="camp-grid">${sorted.map((f, i) => {
    const codes = productCodes(f);
    return `
    <article class="camp-card" role="button" tabindex="0" onclick="openLaunchDetail(${i})" onkeydown="if(event.key==='Enter')openLaunchDetail(${i})">
      <div class="camp-thumb rag-${ragOf(f.Status)}">${ragPill(f.Status)}</div>
      <div class="camp-body">
        <div class="camp-cat">Product launch</div>
        <h3 class="camp-name">${escHtml(f.Title || 'Untitled')}</h3>
        <div class="camp-dates">${escHtml(fmtSpDate(f.LaunchDate) || 'Date to be confirmed')}</div>
        ${codes.length ? `<div class="camp-codes">${escHtml(codes.join(' · '))}</div>` : ''}
      </div>
    </article>`;
  }).join('')}</div>`;
}

let _campaignItems = [];

function renderCampaigns(items) {
  const grid = document.getElementById('sp-campaigns-grid');
  if (!grid) return;

  _campaignItems = items || [];

  // Header metrics — always computed from live data (zeros when empty)
  const count = re => items.filter(f => re.test(String(f.Status || '').toLowerCase())).length;
  const _set = (id, v) => { const n = document.getElementById(id); if (n) n.textContent = v; };
  _set('sp-metric-live',      count(/live/));
  _set('sp-metric-planning',  count(/planning|draft/));
  _set('sp-metric-completed', count(/complete/));
  const totalBudget = items.reduce((sum, f) => sum + (Number(f.Budget) || 0), 0);
  _set('sp-metric-budget', totalBudget ? '£' + totalBudget.toLocaleString('en-GB') : '—');

  if (!items.length) {
    grid.innerHTML = '<p class="prose dim">No campaigns in SharePoint yet — add items to the Campaigns list and they\'ll appear here.</p>';
    return;
  }

  grid.innerHTML = items.map((f, i) => {
    // Same RAG colour code as Product Launches, so a red card means
    // "still planning" on either page.
    const channels = Array.isArray(f.Channels) ? f.Channels.join(' · ') : (f.Channels || '');

    return `
    <article class="camp-card" role="button" tabindex="0" onclick="openCampaignDetail(${i})" onkeydown="if(event.key==='Enter')openCampaignDetail(${i})">
      <div class="camp-thumb rag-${ragOf(f.Status)}">${ragPill(f.Status)}</div>
      <div class="camp-body">
        <div class="camp-cat">${escHtml(f.CampaignType || 'Campaign')}</div>
        <h3 class="camp-name">${escHtml(f.Title || 'Untitled')}</h3>
        <div class="camp-dates">${[fmtSpDate(f.StartDate), fmtSpDate(f.EndDate)].filter(Boolean).join(' – ')}${f.Region ? ' · ' + escHtml(f.Region) : ''}</div>
        <div class="camp-kpis">
          ${f.Budget != null && f.Budget !== '' ? `<div><div class="k">${fmtMoney(f.Budget)}</div><div class="l">Budget</div></div>` : ''}
          ${channels ? `<div><div class="k">${escHtml(channels)}</div><div class="l">Channels</div></div>` : ''}
        </div>
      </div>
    </article>`;
  }).join('');

}

let _docFiles = [];

function renderDocuments(files) {
  const grid = document.getElementById('sp-documents-grid');
  if (!grid) return;

  if (!files.length) {
    grid.innerHTML = '<p class="prose dim">The library is empty — upload files to SharePoint and they\'ll appear here.</p>';
    return;
  }

  _docFiles = files;

  grid.innerHTML = `<div class="asset-grid">${files.map((f, i) => {
    const icon = fileIcon(f.name, !!f.folder);
    const url = safeUrl(f.webUrl, '');
    const meta = [
      f.folder ? `${f.folder.childCount ?? ''} items`.trim() : humanSize(f.size),
      fmtSpDate(f.lastModifiedDateTime),
    ].filter(Boolean).join(' · ');

    // Files open in an in-hub preview modal; folders open in SharePoint.
    const canPreview = !f.folder && f._driveId && f.id;
    const inner = `
      <div class="asset-icon ${icon.cls}">${escHtml(icon.label)}</div>
      <div class="asset-info">
        <div class="asset-name">${escHtml(f.name)}</div>
        <div class="asset-meta">${escHtml(meta)}</div>
      </div>
      ${canPreview ? '<span class="asset-open">Open in hub →</span>' : ''}`;

    return canPreview
      ? `<div class="asset asset-preview" role="button" tabindex="0" onclick="openDocPreview(${i})" onkeydown="if(event.key==='Enter')openDocPreview(${i})">${inner}</div>`
      : `<a class="asset" ${url ? `href="${escAttr(url)}" target="_blank" rel="noopener"` : ''}>${inner}</a>`;
  }).join('')}</div>`;
}

// ── In-hub document preview ───────────────────────────────────
// Uses the Graph "preview" action, which returns a short-lived
// embeddable URL. The file renders in an iframe inside the hub, so
// users never bounce out to SharePoint. Non-previewable types fall
// back to an "Open in SharePoint" link.
// Kept for any legacy callers: preview by index into the last-rendered
// _docFiles array. New code (file browser, campaign blocks) calls
// openDocFile(fileObject) directly.
function openDocPreview(i) {
  return openDocFile(_docFiles[i]);
}

async function openDocFile(f) {
  if (!f) return;

  const modal   = document.getElementById('doc-modal');
  const frame   = document.getElementById('doc-frame');
  const titleEl = document.getElementById('doc-modal-title');
  const spLink  = document.getElementById('doc-modal-splink');
  const loading = document.getElementById('doc-modal-loading');

  if (!modal || !frame) { if (f.webUrl) window.open(f.webUrl, '_blank', 'noopener'); return; }

  if (titleEl) titleEl.textContent = f.name || 'Document';
  if (spLink)  spLink.href = safeUrl(f.webUrl, '#');

  // Download link — the direct URL is short-lived, so fetch it fresh.
  const dl = document.getElementById('doc-modal-download');
  if (dl) {
    dl.style.display = 'none';
    if (f._driveId && f.id) {
      graphFetch(`/drives/${f._driveId}/items/${f.id}?$select=id,name,@microsoft.graph.downloadUrl`)
        .then(meta => {
          const url = meta && meta['@microsoft.graph.downloadUrl'];
          if (!url) return;
          dl.href = url;
          dl.setAttribute('download', f.name || '');
          dl.style.display = '';
        })
        .catch(() => { /* preview still works without it */ });
    }
  }
  const oldFb = document.querySelector('#doc-modal-body .doc-fallback');
  if (oldFb) oldFb.remove();
  frame.removeAttribute('src');
  if (loading) loading.style.display = '';

  modal.classList.remove('hidden');
  document.body.classList.add('modal-open');

  try {
    const prev = await graphPost(`/drives/${f._driveId}/items/${f.id}/preview`, {});
    const url = prev && prev.getUrl;
    if (!url) throw new Error('No preview URL returned');
    frame.onload = () => { if (loading) loading.style.display = 'none'; };
    frame.src = url + (url.includes('?') ? '&' : '?') + 'nb=true';
  } catch (e) {
    if (loading) loading.style.display = 'none';
    const body = document.getElementById('doc-modal-body');
    if (body) body.insertAdjacentHTML('beforeend',
      `<div class="doc-fallback">This file type can't be previewed inline. ` +
      `<a href="${escAttr(safeUrl(f.webUrl, '#'))}" target="_blank" rel="noopener">Open in SharePoint →</a></div>`);
  }
}

function closeDocPreview() {
  const modal = document.getElementById('doc-modal');
  const frame = document.getElementById('doc-frame');
  if (modal) modal.classList.add('hidden');
  if (frame) frame.removeAttribute('src');
  document.body.classList.remove('modal-open');
  const fb = document.querySelector('#doc-modal-body .doc-fallback');
  if (fb) fb.remove();
}

// ═══ Videos — WordPress uploads + SharePoint Media Portal ═════

const VIDEO_EXT = /\.(mp4|mov|m4v|webm)$/i;

async function fetchWordPressVideos() {
  const cached = _cacheGet('videos_wp');
  if (cached) return cached;
  const { apiUrl } = HUB_CONFIG.wordpress;
  const res = await fetch(`${apiUrl}/media?media_type=video&per_page=12&_fields=id,title,source_url,mime_type,date`);
  if (!res.ok) throw new Error(`WordPress media API returned ${res.status}`);
  const items = await res.json();
  const vids = (items || []).filter(v => VIDEO_EXT.test(v.source_url || '')).map(v => ({
    title:  String((v.title && v.title.rendered) || 'Untitled').replace(/&#8217;/g, "'").replace(/&amp;/g, '&').trim(),
    date:   v.date || '',
    src:    v.source_url,   // public CDN mp4 — plays inline
    href:   v.source_url,
    source: 'checkfire.co.uk',
  }));
  _cacheSet('videos_wp', vids);
  return vids;
}

async function fetchSharePointVideos() {
  const cached = _cacheGet('videos_sp');
  if (cached) return cached;

  const u = new URL(HUB_CONFIG.videos.mediaPortalSite);
  const site = await graphFetch(`/sites/${u.hostname}:${u.pathname}`);
  const drives = await graphFetch(`/sites/${site.id}/drives?$select=id,name`);
  const drive = (drives.value || [])[0];
  if (!drive) return [];

  // Graph drive search matches on name fragments — run one query per
  // extension and merge (covers files anywhere in the library,
  // including the "03. Videos" folder).
  const queries = ['mp4', 'mov', 'webm'].map(q =>
    graphFetch(`/drives/${drive.id}/root/search(q='${q}')?$select=name,webUrl,lastModifiedDateTime,file&$top=25`)
      .catch(() => ({ value: [] }))
  );
  const results = await Promise.all(queries);
  const seen = new Set();
  const vids = [];
  for (const r of results) {
    for (const f of (r.value || [])) {
      if (!VIDEO_EXT.test(f.name || '') || seen.has(f.webUrl)) continue;
      seen.add(f.webUrl);
      vids.push({
        title:  f.name.replace(VIDEO_EXT, '').replace(/[-_]+/g, ' ').trim(),
        date:   f.lastModifiedDateTime || '',
        src:    null,          // needs auth — opens in SharePoint's player
        href:   f.webUrl,
        source: 'Media Portal',
      });
    }
  }
  _cacheSet('videos_sp', vids);
  return vids;
}

// YouTube — the channel marketing actually publish to. Read through
// the Azure Function proxy so the YouTube API key stays server-side
// (never put an API key in config.js — see the June 2026 note).
async function fetchYouTubeVideos() {
  const cfg = (HUB_CONFIG.videos && HUB_CONFIG.videos.youtube) || {};
  if (!cfg.proxyUrl) return [];

  const cached = _cacheGet('videos_yt');
  if (cached) return cached;

  const res = await fetch(cfg.proxyUrl, { headers: { Accept: 'application/json' } });
  if (!res.ok) throw new Error('Video proxy returned ' + res.status);
  const data = await res.json();
  const rows = Array.isArray(data) ? data : (data.videos || data.items || []);

  const vids = rows.map(v => {
    const id = v.id || v.videoId || '';
    return {
      title:     String(v.title || 'Untitled').trim(),
      date:      v.date || v.publishedAt || '',
      youtubeId: id,
      src:       null,
      href:      v.url || (id ? 'https://www.youtube.com/watch?v=' + encodeURIComponent(id) : ''),
      thumb:     v.thumb || v.thumbnail || '',
      source:    'YouTube',
    };
  }).filter(v => v.youtubeId);

  _cacheSet('videos_yt', vids);
  return vids;
}

async function loadHomeVideos() {
  const section = document.getElementById('home-videos');
  const grid = document.getElementById('home-videos-grid');
  if (!section || !grid) return;

  const cfg = HUB_CONFIG.videos || {};
  const [yt, wp, sp] = await Promise.allSettled([
    fetchYouTubeVideos(),
    cfg.includeWordPress  === false ? [] : fetchWordPressVideos(),
    cfg.includeSharePoint === false ? [] : fetchSharePointVideos(),
  ]);
  if (yt.status === 'rejected') console.warn('YouTube videos unavailable:', yt.reason.message);
  if (wp.status === 'rejected') console.warn('WordPress videos unavailable:', wp.reason.message);
  if (sp.status === 'rejected') console.warn('SharePoint videos unavailable:', sp.reason.message);

  let vids = [
    ...(yt.status === 'fulfilled' ? yt.value : []),
    ...(wp.status === 'fulfilled' ? wp.value : []),
    ...(sp.status === 'fulfilled' ? sp.value : []),
  ].sort((a, b) => String(b.date).localeCompare(String(a.date)));

  // Only show recent videos (default: last 3 months — see config.js).
  const months = (HUB_CONFIG.videos && HUB_CONFIG.videos.maxAgeMonths) || 0;
  if (months > 0) {
    const cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - months);
    vids = vids.filter(v => v.date && !isNaN(new Date(v.date)) && new Date(v.date) >= cutoff);
  }
  vids = vids.slice(0, (HUB_CONFIG.videos && HUB_CONFIG.videos.max) || 6);

  // Compact hero box: newest 3, links down to the full grid.
  renderHeroVideos(vids);

  if (!vids.length) { section.style.display = 'none'; return; }

  grid.innerHTML = vids.map(v => {
    const href = safeUrl(v.href, '');
    const media = v.youtubeId
      ? `<iframe class="vid-player" src="https://www.youtube-nocookie.com/embed/${escAttr(encodeURIComponent(v.youtubeId))}" title="${escAttr(v.title)}" frameborder="0" loading="lazy" allow="accelerometer; clipboard-write; encrypted-media; picture-in-picture" allowfullscreen></iframe>`
      : v.src
        ? `<video class="vid-player" src="${escAttr(safeUrl(v.src, ''))}" controls preload="metadata" playsinline></video>`
        : `<a class="vid-thumb-link" href="${escAttr(href)}" target="_blank" rel="noopener"><span class="vid-play">▶</span><span>Watch on SharePoint</span></a>`;
    return `
    <div class="vid-card">
      ${media}
      <div class="vid-body">
        <div class="vid-title">${escHtml(v.title)}</div>
        <div class="vid-meta">${escHtml(v.source)}${v.date ? ' · ' + fmtSpDate(v.date) : ''}</div>
      </div>
    </div>`;
  }).join('');
  section.style.display = '';
}

// Compact "Latest Videos" box in the hero — the newest few, each
// scrolling down to the full video grid where they play inline.
function renderHeroVideos(vids) {
  const el = document.getElementById('home-hero-videos-body');
  if (!el) return;
  if (!vids || !vids.length) {
    const ch = safeUrl((HUB_CONFIG.videos && HUB_CONFIG.videos.youtube && HUB_CONFIG.videos.youtube.channelUrl) || '', '');
    el.innerHTML = '<p class="prose dim">No recent videos.</p>' +
      (ch ? `<a class="hbox-more" href="${escAttr(ch)}" target="_blank" rel="noopener">CheckFire on YouTube →</a>` : '');
    return;
  }
  el.innerHTML = vids.slice(0, 3).map(v => `
    <div class="hbox-vid" role="button" tabindex="0" onclick="scrollToVideos()" onkeydown="if(event.key==='Enter')scrollToVideos()">
      <span class="hbox-vid-thumb">▶</span>
      <span class="hbox-vid-title">${escHtml(v.title)}</span>
    </div>`).join('') +
    '<a class="hbox-more" onclick="scrollToVideos()">See all videos →</a>';
}

function scrollToVideos() {
  const s = document.getElementById('home-videos');
  if (s && s.style.display !== 'none') s.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ═══ Team wall ═══════════════════════════════════════════════
// Marketing asked for a proper internal comms wall — post, like,
// comment, tag people, notify everyone. A static site can't do any of
// that on its own, so the wall is a Viva Engage community embedded in
// the hub: it's already in the CheckFire Microsoft 365 licence and all
// of those features come with it, including notifications.
//
// Until the community's embed URL is pasted into config.js, the section
// falls back to a read-only feed of the SharePoint "Comms" list so the
// page still shows something useful.

function _pick(obj, names) {
  for (const n of names) {
    if (obj[n] !== undefined && obj[n] !== null && obj[n] !== '') return obj[n];
  }
  return '';
}

async function fetchCommsItems() {
  const name = (HUB_CONFIG.social && HUB_CONFIG.social.commsList) || 'Comms';
  return fetchListItems(name);
}

function _renderVivaEngage(body) {
  const cfg = (HUB_CONFIG.social || {});
  const embed = safeUrl(cfg.vivaEngageEmbed || '', '');
  // Modern embeds are served from engage.cloud.microsoft; the retired
  // classic ones came from web.yammer.com. Accept either.
  if (!embed || !/^https:\/\/(engage\.cloud\.microsoft|([a-z0-9-]+\.)?yammer\.com)\//i.test(embed)) return false;

  body.innerHTML = `<iframe class="wall-frame" src="${escAttr(embed)}" title="CheckFire team wall" frameborder="0" loading="lazy" allowfullscreen></iframe>`;

  const link = document.getElementById('wall-open-link');
  const open = safeUrl(cfg.vivaEngageUrl || '', '');
  if (link && open) { link.href = open; link.style.display = ''; }
  return true;
}

// Read-only fallback: the Comms list, styled as plain internal
// announcements (no Twitter handles or bird — marketing asked for those
// to go).
function _renderCommsFallback(body, items) {
  let posts = (items || []).map(f => ({
    author: _pick(f, ['Author', 'PostedBy', 'Title']) || 'CheckFire',
    team:   _pick(f, ['Team', 'Department', 'Handle']),
    body:   _pick(f, ['Message', 'Body', 'Post', 'Content', 'Description']),
    date:   _pick(f, ['Date', 'Posted', 'PostDate']) || f.Created || '',
    link:   linkOf(_pick(f, ['Link', 'LinkURL', 'Url'])),
  })).filter(p => p.body);

  posts.sort((a, b) => String(b.date).localeCompare(String(a.date)));
  posts = posts.slice(0, (HUB_CONFIG.social && HUB_CONFIG.social.commsMax) || 8);

  if (!posts.length) {
    body.innerHTML = `
      <div class="wall-empty">
        <h4>The team wall isn't switched on yet</h4>
        <p>Create a Viva Engage community for the team and paste its embed link into <strong>config.js</strong> — that gives everyone posting, likes, comments, @mentions and notifications. In the meantime, anything added to the SharePoint <strong>Comms</strong> list shows up here.</p>
      </div>`;
    return;
  }

  body.innerHTML = `<div class="cm-list wall">${posts.map(p => {
    const init = String(p.author).trim().split(/\s+/).map(w => w[0]).join('').slice(0, 2).toUpperCase() || 'CF';
    const link = safeUrl(p.link, '');
    const meta = [p.team, p.date ? fmtSpDate(p.date) : ''].filter(Boolean).join(' · ');
    const inner = `
      <div class="cm-head">
        <span class="cm-avatar">${escHtml(init)}</span>
        <div class="cm-id">
          <span class="cm-name">${escHtml(p.author)}</span>
          ${meta ? `<span class="cm-handle">${escHtml(meta)}</span>` : ''}
        </div>
      </div>
      <div class="cm-text">${escHtml(p.body)}</div>`;
    return link
      ? `<a class="cm-card" href="${escAttr(link)}" target="_blank" rel="noopener">${inner}</a>`
      : `<div class="cm-card">${inner}</div>`;
  }).join('')}</div>`;
}

async function loadWall() {
  const body = document.getElementById('home-wall-body');
  if (!body) return;

  if (_renderVivaEngage(body)) return;

  try {
    const items = await fetchCommsItems();
    _renderCommsFallback(body, items);
  } catch (e) {
    console.info('[Wall] Comms list not loaded:', e.message);
    _renderCommsFallback(body, []);
  }
}

// ═══ Notices / alerts bar ════════════════════════════════════
// Short "you should know" messages from marketing: a delayed product,
// a website outage, an issue being worked on. Driven by the Notices
// list; the bar stays hidden when there's nothing live.

const _NOTICE_ICONS = {
  info:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" width="18" height="18"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>',
  warning: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" width="18" height="18"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
  alert:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" width="18" height="18"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
};

function _noticeTone(type) {
  const t = String(type || '').toLowerCase();
  if (/alert|urgent|down|critical|outage/.test(t)) return 'alert';
  if (/warn|delay|issue|caution/.test(t))          return 'warning';
  return 'info';
}

function _noticeDismissed(key) {
  try { return sessionStorage.getItem('hubnotice_' + key) === '1'; } catch (_) { return false; }
}

function dismissNotice(key, btn) {
  try { sessionStorage.setItem('hubnotice_' + key, '1'); } catch (_) {}
  const card = btn && btn.closest('.notice');
  if (card) card.remove();
  const wrap = document.getElementById('home-notices');
  if (wrap && !wrap.querySelector('.notice')) wrap.style.display = 'none';
}

let _noticeItems = [];

function renderNotices(items) {
  const wrap = document.getElementById('home-notices');
  if (!wrap) return;

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const isOn = v => v === true || /^(yes|true|1|on)$/i.test(String(v ?? ''));

  let live = (items || []).filter(f => {
    // Active defaults to ON when the column doesn't exist.
    if (f.Active !== undefined && f.Active !== null && f.Active !== '' && !isOn(f.Active)) return false;
    const start = _pick(f, ['StartDate', 'Start']);
    const end   = _pick(f, ['EndDate', 'End', 'Expires']);
    if (start && !isNaN(new Date(start)) && new Date(start) > today) return false;
    if (end   && !isNaN(new Date(end)))   { const e = new Date(end); e.setHours(23, 59, 59); if (e < today) return false; }
    return _pick(f, ['Title', 'Message', 'Body']);
  });

  live.sort((a, b) => String(_pick(b, ['StartDate', 'Created'])).localeCompare(String(_pick(a, ['StartDate', 'Created']))));
  live = live.slice(0, (HUB_CONFIG.notices && HUB_CONFIG.notices.max) || 3);

  _noticeItems = live;

  const cards = live.map((f, i) => {
    const tone  = _noticeTone(_pick(f, ['Type', 'Severity', 'Level']));
    const title = _pick(f, ['Title']);
    const text  = _pick(f, ['Message', 'Body', 'Description']);
    const link  = safeUrl(linkOf(_pick(f, ['Link', 'LinkURL', 'Url'])), '');
    const key   = 'n' + i + '-' + String(title || text).slice(0, 40).replace(/\W+/g, '');
    if (_noticeDismissed(key)) return '';
    return `
      <div class="notice ${tone}">
        <span class="notice-ico">${_NOTICE_ICONS[tone]}</span>
        <div class="notice-body">
          ${title ? `<div class="notice-title">${escHtml(title)}</div>` : ''}
          ${text ? `<div class="notice-text">${escHtml(text)}</div>` : ''}
          ${link ? `<a class="notice-link" href="${escAttr(link)}" target="_blank" rel="noopener">More info →</a>` : ''}
        </div>
        <button class="notice-dismiss" title="Dismiss" onclick="dismissNotice('${escAttr(key)}', this)">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" width="15" height="15"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
      </div>`;
  }).filter(Boolean);

  if (!cards.length) { wrap.style.display = 'none'; wrap.innerHTML = ''; return; }
  wrap.innerHTML = cards.join('');
  wrap.style.display = '';
}

async function loadNotices() {
  const wrap = document.getElementById('home-notices');
  if (!wrap) return;
  try {
    const items = await fetchListItems((HUB_CONFIG.notices && HUB_CONFIG.notices.list) || 'Notices');
    renderNotices(items);
  } catch (e) {
    // No list yet (or not signed in) — stay silent and hidden.
    console.info('[Notices] not loaded:', e.message);
    wrap.style.display = 'none';
  }
}

// ═══ Training calendar ═══════════════════════════════════════
// Internal sessions (e.g. the product-launch training Josh runs) and
// external courses. Dates also feed the marketing calendar as green
// markers — see jotform.js.

let TRAINING_ITEMS = [];

function _trainDate(f) {
  return _pick(f, ['TrainingDate', 'Date', 'StartDate', 'EventDate']);
}

function renderTraining(items) {
  const section = document.getElementById('home-training');
  const track   = document.getElementById('home-training-track');
  if (!section || !track) return;

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const upcoming = (items || [])
    .filter(f => { const d = _trainDate(f); return d && !isNaN(new Date(d)) && new Date(d) >= today; })
    .sort((a, b) => String(_trainDate(a)).localeCompare(String(_trainDate(b))))
    .slice(0, (HUB_CONFIG.training && HUB_CONFIG.training.max) || 8);

  if (!upcoming.length) { section.style.display = 'none'; return; }

  track.innerHTML = upcoming.map(f => {
    const d    = new Date(_trainDate(f));
    const type = String(_pick(f, ['TrainingType', 'Type', 'Category']) || '').toLowerCase();
    const cls  = /extern/.test(type) ? 'external' : 'internal';
    const meta = [
      _pick(f, ['Trainer', 'Host', 'Presenter']),
      _pick(f, ['Location', 'Venue', 'Where']),
    ].filter(Boolean).join(' · ');
    const link = safeUrl(linkOf(_pick(f, ['Link', 'LinkURL', 'Url'])), '');
    const inner = `
      <div class="train-date">
        <div class="d">${d.getDate()}</div>
        <div class="m">${escHtml(d.toLocaleDateString('en-GB', { month: 'short' }))}</div>
      </div>
      <div class="train-info">
        <div class="train-name">${escHtml(_pick(f, ['Title']) || 'Training session')}</div>
        ${meta ? `<div class="train-meta">${escHtml(meta)}</div>` : ''}
        <span class="train-tag ${cls}">${cls === 'external' ? 'External' : 'In-house'}</span>
      </div>`;
    return link
      ? `<a class="train-card" href="${escAttr(link)}" target="_blank" rel="noopener">${inner}</a>`
      : `<div class="train-card">${inner}</div>`;
  }).join('');

  section.style.display = '';
}

async function loadTraining() {
  const section = document.getElementById('home-training');
  if (!section) return;
  try {
    const items = await fetchListItems((HUB_CONFIG.training && HUB_CONFIG.training.list) || 'Training Events');
    TRAINING_ITEMS = items || [];
    renderTraining(TRAINING_ITEMS);
  } catch (e) {
    console.info('[Training] list not loaded:', e.message);
    section.style.display = 'none';
  }
}

// Kept for callers that still ask for "social" — now the wall + training.
async function loadSocial() {
  await Promise.all([loadWall(), loadTraining()]);
}

// ═══ Orchestrators ═══════════════════════════════════════════

function _renderListError(containerId, message, keepExisting) {
  const el = document.getElementById(containerId);
  if (!el) return;
  const html = `<p class="sp-error" style="color:#D1242B;font-size:13px;padding:8px 0">${escHtml(message)}</p>`;
  if (keepExisting) el.insertAdjacentHTML('afterbegin', html);
  else el.innerHTML = html;
}

async function loadSharePointData() {
  if (window.HUB_DEMO_MODE) return;
  if (typeof getAccessToken !== 'function') return;

  const [launches, campaigns] = await Promise.allSettled([
    fetchListItems(HUB_CONFIG.lists.launches),
    fetchListItems(HUB_CONFIG.lists.campaigns),
  ]);

  if (launches.status === 'fulfilled') renderLaunches(launches.value);
  else _renderListError('sp-launches-list', `Couldn't load launches: ${launches.reason.message}`);

  if (campaigns.status === 'fulfilled') renderCampaigns(campaigns.value);
  else _renderListError('sp-campaigns-grid', `Couldn't load campaigns: ${campaigns.reason.message}`, true);

  // Trade & Events is driven by the Documents/Events folders, not a list.
  await loadTradeEvents();
}

// Resources ▸ Marketing Library — opened via loadResourcesData (ui.js).
// Backed by the in-hub file browser so folders open here, not SharePoint.
const _fbLoaded = { marketing: false, product: false };

async function loadSharePointDocuments() {
  const grid = document.getElementById('sp-documents-grid');
  if (!grid) return;
  if (_fbLoaded.marketing) { renderBrowser('marketing'); return; }
  _fbLoaded.marketing = true;
  await fbInit('marketing', HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary,
               'sp-documents-grid', 'docs-crumbs', 'Marketing Library');
}

// Resources ▸ Product Portal tab — the second SharePoint site.
async function loadProductPortal() {
  const grid = document.getElementById('pp-documents-grid');
  if (!grid) return;
  const signedIn = window.AUTH && window.AUTH.account;
  if (window.HUB_DEMO_MODE || !signedIn) {
    grid.innerHTML = '<p class="prose dim">Sign in with your CheckFire account to browse the Product Portal.</p>';
    return;
  }
  if (_fbLoaded.product) { renderBrowser('product'); return; }
  _fbLoaded.product = true;
  await fbInit('product', HUB_CONFIG.productPortalSite, HUB_CONFIG.documentsLibrary,
               'pp-documents-grid', 'pp-crumbs', 'Product Portal');
}

// ═══ In-hub file browser ═════════════════════════════════════
// A small, reusable folder browser. Files open in the in-hub preview
// modal; folders drill in with a breadcrumb trail — the user never
// bounces out to SharePoint. Two instances run independently:
//   'marketing' → Documents library on MarketingHub
//   'product'   → Documents library on the Product Portal site
const FB = {};

function _fbSkeleton(gridId) {
  const g = document.getElementById(gridId);
  if (g) g.innerHTML = '<div class="skeleton sk-line med"></div><div class="skeleton sk-line"></div><div class="skeleton sk-line short"></div>';
}

async function fbInit(key, siteUrl, library, gridId, crumbId, rootLabel, opts) {
  opts = opts || {};
  FB[key] = { siteUrl, library, gridId, crumbId, rootLabel: rootLabel || 'Home',
              driveId: null, rootId: opts.rootItemId || null, path: [], items: [] };
  _fbSkeleton(gridId);
  try {
    const drive = await resolveDrive(siteUrl, library);
    FB[key].driveId = drive.id;
    await fbLoad(key);
  } catch (e) {
    const msg = e.message === 'NOT_FOUND'
      ? 'That SharePoint site or library could not be found — check the URL in config.js and that you have access.'
      : `Couldn't open the library: ${e.message}`;
    _renderListError(gridId, msg);
    _fbLoaded[key] = false;
  }
}

async function fbLoad(key) {
  const b = FB[key];
  if (!b || !b.driveId) return;
  _fbSkeleton(b.gridId);
  const current = b.path.length ? b.path[b.path.length - 1].id : (b.rootId || null);
  try {
    b.items = await fetchDriveChildren(b.driveId, current);
    renderBrowser(key);
  } catch (e) {
    _renderListError(b.gridId, `Couldn't open that folder: ${e.message}`);
  }
}

function renderBrowser(key) {
  const b = FB[key];
  if (!b) return;
  renderCrumbs(key);
  const grid = document.getElementById(b.gridId);
  if (!grid) return;

  if (!b.items.length) {
    grid.innerHTML = '<p class="prose dim">This folder is empty.</p>';
    return;
  }

  // Folders first, then files, each alphabetical.
  const sorted = [...b.items].sort((a, c) =>
    ((c.folder ? 1 : 0) - (a.folder ? 1 : 0)) || String(a.name).localeCompare(String(c.name)));

  grid.innerHTML = `<div class="asset-grid">${sorted.map(f => {
    const idx  = b.items.indexOf(f);
    const icon = fileIcon(f.name, !!f.folder);
    const meta = [
      f.folder ? `${f.folder.childCount ?? ''} items`.trim() : humanSize(f.size),
      fmtSpDate(f.lastModifiedDateTime),
    ].filter(Boolean).join(' · ');
    const inner = `
      <div class="asset-icon ${icon.cls}">${escHtml(icon.label)}</div>
      <div class="asset-info">
        <div class="asset-name">${escHtml(f.name)}</div>
        <div class="asset-meta">${escHtml(meta)}</div>
      </div>
      <span class="asset-open">${f.folder ? 'Open →' : 'Open in hub →'}</span>`;
    if (f.folder) {
      return `<div class="asset asset-preview" role="button" tabindex="0" onclick="fbOpenFolder('${key}',${idx})" onkeydown="if(event.key==='Enter')fbOpenFolder('${key}',${idx})">${inner}</div>`;
    }
    const canPreview = f._driveId && f.id;
    return canPreview
      ? `<div class="asset asset-preview" role="button" tabindex="0" onclick="fbPreview('${key}',${idx})" onkeydown="if(event.key==='Enter')fbPreview('${key}',${idx})">${inner}</div>`
      : `<a class="asset" ${f.webUrl ? `href="${escAttr(safeUrl(f.webUrl))}" target="_blank" rel="noopener"` : ''}>${inner}</a>`;
  }).join('')}</div>`;
}

function renderCrumbs(key) {
  const b = FB[key];
  if (!b) return;
  const el = document.getElementById(b.crumbId);
  if (!el) return;
  const atRoot = b.path.length === 0;
  const parts = [`<span class="fb-crumb${atRoot ? ' current' : ''}" ${atRoot ? '' : `onclick="fbCrumb('${key}',-1)"`}>${escHtml(b.rootLabel)}</span>`];
  b.path.forEach((p, i) => {
    const cur = i === b.path.length - 1;
    parts.push('<span class="fb-sep">/</span>');
    parts.push(`<span class="fb-crumb${cur ? ' current' : ''}" ${cur ? '' : `onclick="fbCrumb('${key}',${i})"`}>${escHtml(p.name)}</span>`);
  });
  el.innerHTML = parts.join('');
}

async function fbOpenFolder(key, idx) {
  const b = FB[key];
  if (!b) return;
  const f = b.items[idx];
  if (!f || !f.folder) return;
  b.path.push({ id: f.id, name: f.name });
  await fbLoad(key);
}

async function fbCrumb(key, i) {
  const b = FB[key];
  if (!b) return;
  b.path = i < 0 ? [] : b.path.slice(0, i + 1);
  await fbLoad(key);
}

function fbPreview(key, idx) {
  const b = FB[key];
  if (!b) return;
  openDocFile(b.items[idx]);
}

// ═══ Campaign / launch detail view ═══════════════════════════
// Clicking a campaign card opens a full detail page (hero, metrics
// bar, and asset blocks). Each asset block maps to a sub-folder inside
//   Documents/Campaigns/<Campaign folder>/<Block folder>
// and opens the file(s) inside — in-hub, never bouncing to SharePoint.

function _num(f, names) {
  for (const n of names) {
    if (f[n] !== undefined && f[n] !== null && f[n] !== '') {
      const v = Number(f[n]);
      if (!isNaN(v)) return v;
    }
  }
  return 0;
}

function _campaignStatusPill(status) {
  const s = String(status || '').toLowerCase();
  if (/live/.test(s))      return '<span class="cd-pill">Live</span>';
  if (/launched/.test(s))  return '<span class="cd-pill">Launched</span>';
  if (/complete/.test(s))  return '<span class="cd-pill done">Completed</span>';
  return `<span class="cd-pill planning">${escHtml(status || 'Planning')}</span>`;
}

// Shared renderer for both Campaigns and Product Launches detail pages.
function _renderDetail(opts) {
  // opts: { containerId, hideIds, item, kind, folderRoot, backLabel, backFn }
  const box = document.getElementById(opts.containerId);
  if (!box) return;
  const f = opts.item || {};

  (opts.hideIds || []).forEach(id => { const el = document.getElementById(id); if (el) el.style.display = 'none'; });
  box.style.display = '';
  window.scrollTo(0, 0);

  const landing = safeUrl(linkOf(f.LinkURL), '');
  const dates = [fmtSpDate(f.StartDate || f.LaunchDate), fmtSpDate(f.EndDate)].filter(Boolean).join(' – ');
  const isLaunch = opts.kind === 'launch';
  // Marketing asked for the launch page to show product codes rather
  // than a marketing blurb.
  const codes = isLaunch ? productCodes(f) : [];
  const sub = isLaunch ? '' : (f.Description || f.Summary || f.CampaignType || '');
  const portal = safeUrl(HUB_CONFIG.productPortalSite || '', '');

  const metrics = opts.kind === 'campaign' ? [
    { label: 'Emails sent',       value: _num(f, ['EmailsSent', 'Emails', 'EmailCount']) },
    { label: 'Social media posts', value: _num(f, ['SocialPosts', 'SocialMediaPosts', 'Social']) },
    { label: 'Blogs',             value: _num(f, ['Blogs', 'BlogCount', 'BlogPosts']) },
    { label: 'PR activity',       value: _num(f, ['PRActivity', 'PR', 'PRActivities']) },
  ] : [];

  const blocks = (HUB_CONFIG.campaignAssetBlocks || []);

  box.innerHTML = `
    <div class="cd-back" role="button" tabindex="0" onclick="${opts.backFn}" onkeydown="if(event.key==='Enter')${opts.backFn}">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><polyline points="15 18 9 12 15 6"/></svg>
      ${escHtml(opts.backLabel)}
    </div>

    <div class="cd-hero">
      <div class="eyebrow" style="color:#D1242B;margin-bottom:8px">${escHtml(opts.kind === 'campaign' ? 'Campaign' : 'Product launch')}</div>
      <h1 class="cd-hero-title">${escHtml(f.Title || 'Untitled')}</h1>
      ${sub ? `<p class="cd-hero-sub">${escHtml(sub)}</p>` : ''}
      ${codes.length ? `<div class="cd-codes">${codes.map(c => `<span class="cd-code">${escHtml(c)}</span>`).join('')}</div>` : ''}
      <div class="cd-hero-meta">
        ${ragChip(f.Status)}
        ${dates ? `<span>${escHtml(dates)}</span>` : ''}
        ${f.Region ? `<span>${escHtml(f.Region)}</span>` : ''}
        ${landing ? `<a class="cd-landing" href="${escAttr(landing)}" target="_blank" rel="noopener">View landing page →</a>` : ''}
        ${isLaunch && portal ? `<a class="cd-portal" href="${escAttr(portal)}" target="_blank" rel="noopener">Open in Product Portal →</a>` : ''}
      </div>
    </div>

    ${metrics.length ? `<div class="cd-metrics">${metrics.map(m => `
      <div class="cd-metric"><div class="cd-metric-label">${escHtml(m.label)}</div><div class="cd-metric-value">${m.value}</div></div>
    `).join('')}</div>` : ''}

    <p class="cd-blocks-title">Assets &amp; resources</p>
    <div class="cd-blocks" id="cd-blocks">${_blocksHtml(blocks)}</div>

    <div id="cd-asset-panel" style="margin-top:20px"></div>`;

  // Remember what the asset blocks should resolve against.
  _detailContext = { folderRoot: opts.folderRoot, campaignFolder: f.CampaignFolder || f.Folder || f.Title };
  _detailBlocks  = blocks;

  // Then swap the config placeholders for the real SharePoint folders.
  _loadDetailBlocks();
}

// Render the asset-block tiles. Each tile shows how many files are in
// its folder once we've read the library.
function _blocksHtml(blocks) {
  return (blocks || []).map((bl, bi) => `
      <div class="cd-block" role="button" tabindex="0" onclick="openDetailAsset(${bi})" onkeydown="if(event.key==='Enter')openDetailAsset(${bi})">
        <span class="cd-block-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="22" height="22"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg></span>
        ${escHtml(bl.label)}
        <span class="cd-block-note">${bl.count === undefined ? 'Open in hub'
          : `${bl.count} item${bl.count === 1 ? '' : 's'} · Open in hub`}</span>
      </div>`).join('');
}

// Read the real sub-folders under Documents/<root>/<item>/ so the tiles
// match what's actually in SharePoint instead of a fixed list in
// config.js. Falls back silently to the config blocks.
async function _loadDetailBlocks() {
  const ctx = _detailContext;
  const box = document.getElementById('cd-blocks');
  if (!ctx || !box) return;

  try {
    const drive = await resolveDrive(HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary);
    const rootFolder = await _findChildFolder(drive.id, null, ctx.folderRoot);
    if (!rootFolder) return;
    const itemFolder = await _findChildFolder(drive.id, rootFolder.id, ctx.campaignFolder);
    if (!itemFolder) return;

    const kids = (await fetchDriveChildren(drive.id, itemFolder.id)).filter(x => x.folder);
    if (!kids.length) return;

    _detailBlocks = kids
      .sort((a, b) => String(a.name).localeCompare(String(b.name)))
      .map(k => ({ label: k.name, folder: k.name, count: (k.folder && k.folder.childCount) || 0 }));

    // Guard against the user having navigated away while we were loading.
    if (_detailContext === ctx && document.getElementById('cd-blocks') === box) {
      box.innerHTML = _blocksHtml(_detailBlocks);
    }
  } catch (e) {
    console.info('[Assets] using the config block list:', e.message);
  }
}

let _detailBlocks = [];

let _detailContext = null;

function openCampaignDetail(i) {
  const f = _campaignItems[i];
  if (!f) return;
  _renderDetail({
    containerId: 'campaign-detail',
    hideIds: ['campaigns-head', 'campaigns-list'],
    item: f,
    kind: 'campaign',
    folderRoot: (HUB_CONFIG.folders && HUB_CONFIG.folders.campaigns) || 'Campaigns',
    backLabel: 'Back to campaigns',
    backFn: 'closeCampaignDetail()',
  });
}

function closeCampaignDetail() {
  ['campaigns-head', 'campaigns-list'].forEach(id => { const el = document.getElementById(id); if (el) el.style.display = ''; });
  const box = document.getElementById('campaign-detail');
  if (box) { box.style.display = 'none'; box.innerHTML = ''; }
  _detailContext = null;
  window.scrollTo(0, 0);
}

function openLaunchDetail(i) {
  const f = _launchItems[i];
  if (!f) return;
  _renderDetail({
    containerId: 'launch-detail',
    hideIds: ['launches-head', 'launches-list-wrap'],
    item: f,
    kind: 'launch',
    folderRoot: (HUB_CONFIG.folders && HUB_CONFIG.folders.launches) || 'Launches',
    backLabel: 'Back to launches',
    backFn: 'closeLaunchDetail()',
  });
}

function closeLaunchDetail() {
  ['launches-head', 'launches-list-wrap'].forEach(id => { const el = document.getElementById(id); if (el) el.style.display = ''; });
  const box = document.getElementById('launch-detail');
  if (box) { box.style.display = 'none'; box.innerHTML = ''; }
  _detailContext = null;
  window.scrollTo(0, 0);
}

// Find a child folder by (case-insensitive) name under a parent item.
async function _findChildFolder(driveId, parentId, name) {
  const target = String(name || '').trim().toLowerCase();
  if (!target) return null;
  const items = await fetchDriveChildren(driveId, parentId);
  return items.find(f => f.folder && String(f.name).trim().toLowerCase() === target)
      || items.find(f => f.folder && String(f.name).trim().toLowerCase().includes(target));
}

// Resolve  Documents/<folderRoot>/<campaignFolder>/<block.folder>  and open
// its file(s) in-hub. One file opens straight into the preview; several are
// listed in a panel; none shows a friendly "not set up yet" note.
async function openDetailAsset(blockIdx) {
  const ctx   = _detailContext;
  const block = (_detailBlocks && _detailBlocks.length ? _detailBlocks : (HUB_CONFIG.campaignAssetBlocks || []))[blockIdx];
  const panel = document.getElementById('cd-asset-panel');
  if (!ctx || !block || !panel) return;

  panel.innerHTML = `<p class="prose dim">Opening “${escHtml(block.label)}”…</p>`;

  try {
    const drive = await resolveDrive(HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary);
    const rootFolder = await _findChildFolder(drive.id, null, ctx.folderRoot);
    if (!rootFolder) throw new Error(`No “${ctx.folderRoot}” folder in the document library yet.`);
    const itemFolder = await _findChildFolder(drive.id, rootFolder.id, ctx.campaignFolder);
    if (!itemFolder) throw new Error(`No folder named “${ctx.campaignFolder}” inside ${ctx.folderRoot} yet.`);
    const blockFolder = await _findChildFolder(drive.id, itemFolder.id, block.folder);
    if (!blockFolder) throw new Error(`No “${block.label}” folder set up for this item yet.`);

    const files = (await fetchDriveChildren(drive.id, blockFolder.id)).filter(x => !x.folder);
    if (!files.length) { panel.innerHTML = `<p class="prose dim">No files in “${escHtml(block.label)}” yet.</p>`; return; }

    if (files.length === 1) {
      _lastAssetFiles = files;
      openDocFile(files[0]);
      panel.innerHTML = `<p class="prose dim">Opened <strong>${escHtml(files[0].name)}</strong> — <a class="fb-crumb" onclick="openDocFile(_lastAssetFiles[0])">reopen</a></p>`;
      return;
    }

    _lastAssetFiles = files;
    panel.innerHTML = `
      <p class="cd-blocks-title" style="margin-bottom:10px">${escHtml(block.label)} — ${files.length} files</p>
      <div class="asset-grid">${files.map((f, idx) => {
        const icon = fileIcon(f.name, false);
        const meta = [humanSize(f.size), fmtSpDate(f.lastModifiedDateTime)].filter(Boolean).join(' · ');
        return `<div class="asset asset-preview" role="button" tabindex="0" onclick="openDocFile(_lastAssetFiles[${idx}])" onkeydown="if(event.key==='Enter')openDocFile(_lastAssetFiles[${idx}])">
          <div class="asset-icon ${icon.cls}">${escHtml(icon.label)}</div>
          <div class="asset-info"><div class="asset-name">${escHtml(f.name)}</div><div class="asset-meta">${escHtml(meta)}</div></div>
          <span class="asset-open">Open in hub →</span>
        </div>`;
      }).join('')}</div>`;
  } catch (e) {
    panel.innerHTML = `<p class="prose dim">${escHtml(e.message)}</p>`;
  }
}


// ═══ Trade & customer events ═════════════════════════════════
// Everything on this page comes from the "Events" folder in the
// marketing document library. Each sub-folder is one event (e.g.
// "FSE 2027") and holds the forms the sales team need beforehand.
// A four-digit year in the folder name splits upcoming from previous.

let _eventFolders = [];

function _eventYear(name) {
  const m = String(name || '').match(/(19|20)\d{2}/);
  return m ? parseInt(m[0], 10) : null;
}

async function loadTradeEvents() {
  const upWrap   = document.getElementById('ev-upcoming-grid');
  const pastWrap = document.getElementById('ev-past-grid');
  if (!upWrap || !pastWrap) return;

  const _set = (id, v) => { const n = document.getElementById(id); if (n) n.textContent = v; };

  try {
    const drive  = await resolveDrive(HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary);
    const rootNm = (HUB_CONFIG.tradeEvents && HUB_CONFIG.tradeEvents.folder) || 'Events';
    const root   = await _findChildFolder(drive.id, null, rootNm);
    if (!root) throw new Error(`No "${rootNm}" folder in the document library yet.`);

    const kids = (await fetchDriveChildren(drive.id, root.id)).filter(x => x.folder);
    _eventFolders = kids.map(k => ({
      id:    k.id,
      name:  k.name,
      year:  _eventYear(k.name),
      count: (k.folder && k.folder.childCount) || 0,
      modified: k.lastModifiedDateTime,
      driveId: drive.id,
    }));

    const thisYear = new Date().getFullYear();
    // No year in the name? Treat it as current/ongoing.
    const upcoming = _eventFolders.filter(e => e.year === null || e.year >= thisYear)
      .sort((a, b) => (a.year || thisYear) - (b.year || thisYear) || String(a.name).localeCompare(String(b.name)));
    const past = _eventFolders.filter(e => e.year !== null && e.year < thisYear)
      .sort((a, b) => b.year - a.year || String(a.name).localeCompare(String(b.name)));

    _set('sp-metric-ev-upcoming', upcoming.length);
    _set('sp-metric-ev-past',     past.length);

    const card = (e, isPast) => {
      const idx = _eventFolders.indexOf(e);
      return `
        <div class="ev-card${isPast ? ' past' : ''}" role="button" tabindex="0" onclick="openEventFolder(${idx})" onkeydown="if(event.key==='Enter')openEventFolder(${idx})">
          <div class="ev-year">${escHtml(e.year ? String(e.year) : 'Ongoing')}</div>
          <div class="ev-name">${escHtml(e.name)}</div>
          <div class="ev-meta">${e.count} item${e.count === 1 ? '' : 's'} · Open in hub →</div>
        </div>`;
    };

    upWrap.innerHTML = upcoming.length
      ? upcoming.map(e => card(e, false)).join('')
      : '<p class="prose dim">No upcoming events yet — add a folder under Documents ▸ Events.</p>';
    pastWrap.innerHTML = past.length
      ? past.map(e => card(e, true)).join('')
      : '<p class="prose dim">Nothing archived yet.</p>';

  } catch (e) {
    _set('sp-metric-ev-upcoming', '–');
    _set('sp-metric-ev-past', '–');
    const msg = e.message === 'NOT_FOUND'
      ? 'Could not reach the document library — check you have access to the MarketingHub site.'
      : e.message;
    _renderListError('ev-upcoming-grid', msg);
    pastWrap.innerHTML = '';
  }
}

function openEventFolder(idx) {
  const ev = _eventFolders[idx];
  if (!ev) return;

  const overview = document.getElementById('events-overview');
  const browser  = document.getElementById('events-browser');
  const title    = document.getElementById('ev-browser-title');
  if (overview) overview.style.display = 'none';
  if (browser)  browser.style.display  = '';
  if (title)    title.textContent = ev.name;
  window.scrollTo(0, 0);

  fbInit('events', HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary,
         'ev-documents-grid', 'ev-crumbs', ev.name, { rootItemId: ev.id });
}

function closeEventFolder() {
  const overview = document.getElementById('events-overview');
  const browser  = document.getElementById('events-browser');
  if (overview) overview.style.display = '';
  if (browser)  browser.style.display  = 'none';
  window.scrollTo(0, 0);
}

let _lastAssetFiles = [];

// ═══ Polls ═══════════════════════════════════════════════════
// A live poll card on the home page. Marketing writes the question in
// the SharePoint "Polls" list; everyone answers with one click IN THE
// HUB, and the click writes a row to "Poll Votes". One vote per person
// per poll. See the HUB_CONFIG.polls comment block in config.js for the
// two lists and their columns.
//
// Reading uses the hub's normal read-only token. Writing (i.e. voting)
// asks for Sites.ReadWrite.All the first time — see getWriteToken() in
// auth.js. If that is refused the card degrades to a read-only result.

const POLL = { poll: null, votes: [], busy: false };

// Like fetchListItems() but keeps the SharePoint item id, which the
// poll needs in order to tie votes to a question. Deliberately uncached
// so a vote shows up straight away.
async function _fetchListRows(listName, top) {
  const siteId = await getSiteId();
  const data = await graphFetch(
    `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items?expand=fields&$top=${top || 100}`
  );
  return (data.value || []).map(i => ({ id: i.id, fields: i.fields || {} }));
}

// "Options" is a multi-line column — one option per line. If the column
// was created as rich text it comes back as HTML, so tags become line
// breaks and the usual entities are decoded.
function _pollOptions(f) {
  const raw = String(f.Options || f.options || f.Choices || '');
  return raw
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li)>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .split(/\r?\n/)
    .map(x => x.trim())
    .filter(Boolean)
    .slice(0, 8);
}

function _pollIsLive(f) {
  const a = f.Active;
  if (a === false || a === 0 || a === 'No' || a === 'false') return false;
  if (f.EndDate) {
    const d = new Date(f.EndDate);
    const today = new Date(); today.setHours(0, 0, 0, 0);
    if (!isNaN(d) && d < today) return false;
  }
  return true;
}

function _pollMyEmail() {
  return String((window.AUTH && window.AUTH.account && window.AUTH.account.mail) || '').toLowerCase();
}

function _pollNote(msg) {
  const el = document.getElementById('poll-note');
  if (el) el.innerHTML = msg ? escHtml(msg) : '';
}

async function loadPolls() {
  const card = document.getElementById('home-poll');
  const body = document.getElementById('home-poll-body');
  if (!card || !body) return;

  const cfg = HUB_CONFIG.polls || {};
  const signedIn = window.AUTH && window.AUTH.account;
  if (window.HUB_DEMO_MODE || !signedIn) {
    body.innerHTML = '<p class="poll-empty">Sign in to see the current poll.</p>';
    return;
  }

  try {
    const rows = await _fetchListRows(cfg.list || 'Polls');
    const live = rows.filter(r => _pollIsLive(r.fields) && _pollOptions(r.fields).length >= 2);
    // Newest first — SharePoint item ids increase.
    const poll = live.sort((a, b) => Number(b.id) - Number(a.id))[0];

    if (!poll) {
      POLL.poll = null;
      body.innerHTML =
        '<p class="poll-empty">No poll running right now.<br>' +
        '<span class="poll-hint">Add a question to the <strong>Polls</strong> list to start one.</span></p>';
      return;
    }

    POLL.poll = poll;
    try {
      const votes = await _fetchListRows(cfg.votesList || 'Poll Votes', 500);
      POLL.votes = votes.filter(v => String(v.fields.PollId || '') === String(poll.id));
    } catch (e) {
      POLL.votes = [];
      console.info('[Polls] votes list not readable yet:', e.message);
    }
    renderPoll();
  } catch (e) {
    console.info('[Polls] not loaded:', e.message);
    body.innerHTML = '<p class="poll-empty">Poll unavailable.<br><span class="poll-hint">Create a <strong>Polls</strong> list on the MarketingHub site.</span></p>';
  }
}

function renderPoll() {
  const body = document.getElementById('home-poll-body');
  if (!body || !POLL.poll) return;

  const f    = POLL.poll.fields;
  const opts = _pollOptions(f);
  const me   = _pollMyEmail();

  const mine  = POLL.votes.find(v => String(v.fields.Voter || '').toLowerCase() === me);
  const total = POLL.votes.length;

  const counts = {};
  opts.forEach(o => { counts[o] = 0; });
  POLL.votes.forEach(v => {
    const t = String(v.fields.Title || '');
    if (counts[t] !== undefined) counts[t]++;
  });

  const question = escHtml(f.Title || 'Quick poll');

  let inner;
  if (mine) {
    const chosen = String(mine.fields.Title || '');
    inner = opts.map(o => {
      const n   = counts[o] || 0;
      const pct = total ? Math.round((n / total) * 100) : 0;
      return `
        <div class="poll-res${o === chosen ? ' mine' : ''}">
          <div class="poll-res-top">
            <span class="poll-res-lbl">${escHtml(o)}${o === chosen ? ' <span class="poll-tick">✓</span>' : ''}</span>
            <span class="poll-res-pct">${pct}%</span>
          </div>
          <div class="poll-bar"><span style="width:${pct}%"></span></div>
        </div>`;
    }).join('');
  } else {
    inner = opts.map((o, i) => `
      <button type="button" class="poll-opt" onclick="votePoll(${i})">${escHtml(o)}</button>
    `).join('');
  }

  body.innerHTML = `
    <div class="poll-q">${question}</div>
    <div class="poll-opts">${inner}</div>
    <div class="poll-foot">
      <span>${total} vote${total === 1 ? '' : 's'}${mine ? ' · thanks for voting' : ''}</span>
      <span id="poll-note" class="poll-note"></span>
    </div>`;
}

async function votePoll(i) {
  if (POLL.busy || !POLL.poll) return;
  const opts   = _pollOptions(POLL.poll.fields);
  const choice = opts[i];
  if (!choice) return;

  POLL.busy = true;
  _pollNote('Saving…');

  try {
    const token = (typeof getWriteToken === 'function') ? await getWriteToken() : null;
    if (!token) {
      _pollNote('Needs permission to save — ask IT to approve hub write access.');
      POLL.busy = false;
      return;
    }

    const siteId   = await getSiteId();
    const listName = (HUB_CONFIG.polls && HUB_CONFIG.polls.votesList) || 'Poll Votes';
    const res = await fetch(
      `${GRAPH_BASE}/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`,
      {
        method:  'POST',
        headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fields: {
            Title:  choice,
            PollId: String(POLL.poll.id),
            Voter:  _pollMyEmail(),
          },
        }),
      }
    );
    if (!res.ok) throw new Error('SharePoint returned ' + res.status);

    const created = await res.json();
    POLL.votes.push({
      id: created.id,
      fields: created.fields || { Title: choice, PollId: String(POLL.poll.id), Voter: _pollMyEmail() },
    });
    renderPoll();
  } catch (e) {
    _pollNote('Not saved — ' + e.message);
  }
  POLL.busy = false;
}
