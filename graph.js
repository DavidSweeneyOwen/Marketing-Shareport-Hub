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

// ═══ Renderers ═══════════════════════════════════════════════

let _launchItems = [];

function renderLaunches(items) {
  const el = document.getElementById('sp-launches-list');
  if (!el) return;

  if (!items.length) {
    el.innerHTML = '<p class="prose dim">No launches in SharePoint yet — add items to the Product Launches list.</p>';
    return;
  }

  const sorted = [...items].sort((a, b) => String(a.LaunchDate || '').localeCompare(String(b.LaunchDate || '')));
  _launchItems = sorted;

  el.innerHTML = `<div class="asset-grid">${sorted.map((f, i) => {
    return `
    <div class="asset asset-preview" role="button" tabindex="0" onclick="openLaunchDetail(${i})" onkeydown="if(event.key==='Enter')openLaunchDetail(${i})">
      <div class="asset-info">
        <div class="asset-name">${escHtml(f.Title || 'Untitled')}</div>
        <div class="asset-meta">${[
          escHtml(f.SKU || ''),
          fmtSpDate(f.LaunchDate),
          f.RRP != null && f.RRP !== '' ? fmtMoney(f.RRP) : '',
        ].filter(Boolean).join(' · ')}</div>
      </div>
      ${statusBadge(f.Status)}
      <span class="asset-open">Open →</span>
    </div>`;
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

  const tones = ['tone-red', 'tone-teal', 'tone-amber', 'tone-blue', 'tone-gold', 'tone-slate'];

  grid.innerHTML = items.map((f, i) => {
    const s = String(f.Status || '').toLowerCase();
    const pill = /live/.test(s)
      ? '<span class="pill live"><span class="status-dot green"></span>Live</span>'
      : /complete/.test(s)
        ? '<span class="pill done">Completed</span>'
        : `<span class="pill planning">${escHtml(f.Status || 'Planning')}</span>`;
    const channels = Array.isArray(f.Channels) ? f.Channels.join(' · ') : (f.Channels || '');

    return `
    <article class="camp-card" role="button" tabindex="0" onclick="openCampaignDetail(${i})" onkeydown="if(event.key==='Enter')openCampaignDetail(${i})">
      <div class="camp-thumb ${tones[i % tones.length]}">${pill}</div>
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

function renderEvents(items) {
  const el = document.getElementById('sp-events-list');
  if (!el) return;

  // Header metrics — always computed from live data
  const count = re => items.filter(f => re.test(String(f.Status || '').toLowerCase())).length;
  const _set = (id, v) => { const n = document.getElementById(id); if (n) n.textContent = v; };
  _set('sp-metric-shows-confirmed', count(/confirmed|live/));
  _set('sp-metric-shows-planning',  count(/planning|draft|upcoming/));

  if (!items.length) {
    el.innerHTML = '<p class="prose dim">No events in SharePoint yet — add items to the Events list and they\'ll appear here.</p>';
    return;
  }

  const sorted = [...items].sort((a, b) => String(a.EventDate || '').localeCompare(String(b.EventDate || '')));

  el.innerHTML = sorted.map(f => {
    const d = f.EventDate ? new Date(f.EventDate) : null;
    const day = d && !isNaN(d) ? String(d.getDate()).padStart(2, '0') : '—';
    const mon = d && !isNaN(d) ? d.toLocaleDateString('en-GB', { month: 'short', year: 'numeric' }) : '';
    const isUpcoming = d && d >= new Date();
    const url = safeUrl(linkOf(f.LinkURL), '');

    return `
    <article class="show" ${url ? `onclick="window.open('${escAttr(url)}','_blank','noopener')"` : ''}>
      <div class="show-date ${isUpcoming ? 'upcoming' : ''}">
        <div class="d">${day}</div><div class="m">${escHtml(mon)}</div>
      </div>
      <div class="show-body">
        <div class="show-top">
          <div>
            <h3 class="show-name">${escHtml(f.Title || 'Untitled')}</h3>
            <div class="show-where">${escHtml([f.Location, f.EventType].filter(Boolean).join(' · '))}</div>
          </div>
          ${statusBadge(f.Status)}
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

  // Direct download (e.g. email signatures, images for Outlook). The
  // short-lived download URL is fetched fresh each time the viewer opens.
  const dl = document.getElementById('doc-modal-download');
  if (dl) {
    dl.style.display = 'none';
    dl.removeAttribute('href');
    if (f._driveId && f.id) {
      graphFetch(`/drives/${f._driveId}/items/${f.id}?$select=id,name,@microsoft.graph.downloadUrl`)
        .then(it => {
          const u = it && it['@microsoft.graph.downloadUrl'];
          if (u) { dl.href = u; dl.setAttribute('download', f.name || ''); dl.style.display = ''; }
        }).catch(() => {});
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

async function loadHomeVideos() {
  const section = document.getElementById('home-videos');
  const grid = document.getElementById('home-videos-grid');
  if (!section || !grid) return;

  const [wp, sp] = await Promise.allSettled([fetchWordPressVideos(), fetchSharePointVideos()]);
  if (wp.status === 'rejected') console.warn('WordPress videos unavailable:', wp.reason.message);
  if (sp.status === 'rejected') console.warn('SharePoint videos unavailable:', sp.reason.message);

  let vids = [
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
    const media = v.src
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
    el.innerHTML = '<p class="prose dim">No recent videos.</p>' +
      '<a class="hbox-more" href="https://checkfireltd.sharepoint.com/sites/CheckFireMediaPortal" target="_blank" rel="noopener">Media Portal →</a>';
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

// ═══ Social — LinkedIn embeds + in-house comms feed ══════════

function _pick(obj, names) {
  for (const n of names) {
    if (obj[n] !== undefined && obj[n] !== null && obj[n] !== '') return obj[n];
  }
  return '';
}

// LinkedIn has no free page feed — we embed specific posts whose embed
// URLs are pasted into HUB_CONFIG.social.linkedInEmbeds. Panel hides
// itself when none are configured.
function renderLinkedIn() {
  const wrap = document.getElementById('home-linkedin');
  const list = document.getElementById('home-linkedin-list');
  if (!wrap || !list) return;

  const embeds = (HUB_CONFIG.social && HUB_CONFIG.social.linkedInEmbeds) || [];
  const frames = embeds
    .map(src => safeUrl(src, ''))
    .filter(u => u && /(^https:\/\/)([a-z]+\.)?linkedin\.com\//i.test(u))
    .map(u => `<iframe class="li-embed" src="${escAttr(u)}" height="430" frameborder="0" allowfullscreen title="Embedded LinkedIn post" loading="lazy"></iframe>`);

  if (!frames.length) { wrap.style.display = 'none'; return; }
  list.innerHTML = frames.join('');
  wrap.style.display = '';
}

// In-house comms: a Twitter-style stream driven by a SharePoint list
// on the MarketingHub site (HUB_CONFIG.social.commsList). Columns are
// read defensively so minor naming differences still work.
async function fetchCommsItems() {
  const name = (HUB_CONFIG.social && HUB_CONFIG.social.commsList) || 'Comms';
  return fetchListItems(name);
}

function renderComms(items) {
  const wrap = document.getElementById('home-comms');
  const list = document.getElementById('home-comms-list');
  if (!wrap || !list) return;

  let posts = (items || []).map(f => ({
    author: _pick(f, ['Author', 'PostedBy', 'Title']) || 'CheckFire',
    handle: _pick(f, ['Handle', 'Team', 'Department']),
    body:   _pick(f, ['Message', 'Body', 'Post', 'Content', 'Description']),
    date:   _pick(f, ['Date', 'Posted', 'PostDate']) || f.Created || '',
    link:   linkOf(_pick(f, ['Link', 'LinkURL', 'Url'])),
  })).filter(p => p.body);

  posts.sort((a, b) => String(b.date).localeCompare(String(a.date)));
  posts = posts.slice(0, (HUB_CONFIG.social && HUB_CONFIG.social.commsMax) || 8);

  if (!posts.length) { wrap.style.display = 'none'; return; }

  const bird = `<svg class="cm-bird" viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M23 4.9c-.8.4-1.7.6-2.6.8a4.5 4.5 0 0 0 2-2.5c-.9.5-1.8.9-2.9 1.1a4.5 4.5 0 0 0-7.7 4.1A12.8 12.8 0 0 1 2.7 3.6a4.5 4.5 0 0 0 1.4 6 4.4 4.4 0 0 1-2-.6v.1a4.5 4.5 0 0 0 3.6 4.4 4.5 4.5 0 0 1-2 .1 4.5 4.5 0 0 0 4.2 3.1A9 9 0 0 1 1 21.5a12.7 12.7 0 0 0 6.9 2c8.3 0 12.8-6.9 12.8-12.8v-.6c.9-.6 1.6-1.4 2.3-2.2z"/></svg>`;

  list.innerHTML = posts.map(p => {
    const handleTxt = p.handle ? '@' + String(p.handle).replace(/\s+/g, '').toLowerCase() : '@checkfire';
    const init = String(p.author).trim().split(/\s+/).map(w => w[0]).join('').slice(0, 2).toUpperCase() || 'CF';
    const link = safeUrl(p.link, '');
    const inner = `
      <div class="cm-head">
        <span class="cm-avatar">${escHtml(init)}</span>
        <div class="cm-id">
          <span class="cm-name">${escHtml(p.author)}</span>
          <span class="cm-handle">${escHtml(handleTxt)}</span>
        </div>
        ${bird}
      </div>
      <div class="cm-text">${escHtml(p.body)}</div>
      ${p.date ? `<div class="cm-date">${escHtml(fmtSpDate(p.date))}</div>` : ''}`;
    return link
      ? `<a class="cm-card" href="${escAttr(link)}" target="_blank" rel="noopener">${inner}</a>`
      : `<div class="cm-card">${inner}</div>`;
  }).join('');
  wrap.style.display = '';
}

async function loadSocial() {
  renderLinkedIn();
  try {
    const items = await fetchCommsItems();
    renderComms(items);
  } catch (e) {
    console.info('[Comms] feed not loaded:', e.message);
    const wrap = document.getElementById('home-comms');
    if (wrap) wrap.style.display = 'none';
  }
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

  const [launches, campaigns, events] = await Promise.allSettled([
    fetchListItems(HUB_CONFIG.lists.launches),
    fetchListItems(HUB_CONFIG.lists.campaigns),
    fetchListItems(HUB_CONFIG.lists.events),
  ]);

  if (launches.status === 'fulfilled') renderLaunches(launches.value);
  else _renderListError('sp-launches-list', `Couldn't load launches: ${launches.reason.message}`);

  if (campaigns.status === 'fulfilled') renderCampaigns(campaigns.value);
  else _renderListError('sp-campaigns-grid', `Couldn't load campaigns: ${campaigns.reason.message}`, true);

  if (events.status === 'fulfilled') renderEvents(events.value);
  else _renderListError('sp-events-list', `Couldn't load events: ${events.reason.message}`, true);
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
  if (_fbLoaded.product) { renderBrowser('product'); _ppRenderSwitch(); return; }
  _fbLoaded.product = true;
  await fbInit('product', HUB_CONFIG.productPortalSite, _ppCurrentLib(),
               'pp-documents-grid', 'pp-crumbs', 'Product Portal');
  _ppRenderSwitch();
}

// The Product Portal keeps files in more than one library (e.g. "Data
// Sheets" and "Documents"). Pills above the browser switch between them.
function _ppLibs() {
  const libs = HUB_CONFIG.productPortalLibraries;
  return (Array.isArray(libs) && libs.length) ? libs : [HUB_CONFIG.documentsLibrary];
}
function _ppCurrentLib() {
  return window._ppLib || _ppLibs()[0];
}
function _ppRenderSwitch() {
  const libs = _ppLibs();
  if (libs.length < 2) return;
  const crumbs = document.getElementById('pp-crumbs');
  if (!crumbs) return;
  let bar = document.getElementById('pp-lib-switch');
  if (!bar) {
    bar = document.createElement('div');
    bar.id = 'pp-lib-switch';
    bar.style.cssText = 'display:flex;gap:8px;margin:0 0 12px';
    crumbs.parentNode.insertBefore(bar, crumbs);
  }
  const cur = _ppCurrentLib();
  bar.innerHTML = libs.map(l => {
    const on = l === cur;
    const style = on
      ? 'background:#111;color:#fff;border:1px solid #111'
      : 'background:#fff;color:#111;border:1px solid #DADADA';
    return `<button type="button" style="${style};font-size:12px;font-weight:700;border-radius:20px;padding:6px 14px;cursor:pointer" onclick="ppSwitchLib('${escAttr(l)}')">${escHtml(l)}</button>`;
  }).join('');
}
async function ppSwitchLib(lib) {
  if (lib === _ppCurrentLib()) return;
  window._ppLib = lib;
  delete FB.product;
  await fbInit('product', HUB_CONFIG.productPortalSite, lib,
               'pp-documents-grid', 'pp-crumbs', 'Product Portal');
  _ppRenderSwitch();
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

async function fbInit(key, siteUrl, library, gridId, crumbId, rootLabel) {
  FB[key] = { siteUrl, library, gridId, crumbId, rootLabel: rootLabel || 'Home', driveId: null, path: [], items: [] };
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
  const current = b.path.length ? b.path[b.path.length - 1].id : null;
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
  const sub = f.Description || f.Summary || f.CampaignType || '';

  const metrics = opts.kind === 'campaign' ? [
    { label: 'Emails sent',       value: _num(f, ['EmailsSent', 'Emails', 'EmailCount']) },
    { label: 'Social media posts', value: _num(f, ['SocialPosts', 'SocialMediaPosts', 'Social']) },
    { label: 'Blogs',             value: _num(f, ['Blogs', 'BlogCount', 'BlogPosts']) },
    { label: 'PR activity',       value: _num(f, ['PRActivity', 'PR', 'PRActivities']) },
  ] : [];

  box.innerHTML = `
    <div class="cd-back" role="button" tabindex="0" onclick="${opts.backFn}" onkeydown="if(event.key==='Enter')${opts.backFn}">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><polyline points="15 18 9 12 15 6"/></svg>
      ${escHtml(opts.backLabel)}
    </div>

    <div class="cd-hero">
      <div class="eyebrow" style="color:#D1242B;margin-bottom:8px">${escHtml(opts.kind === 'campaign' ? 'Campaign' : 'Product launch')}</div>
      <h1 class="cd-hero-title">${escHtml(f.Title || 'Untitled')}</h1>
      ${sub ? `<p class="cd-hero-sub">${escHtml(sub)}</p>` : ''}
      <div class="cd-hero-meta">
        ${_campaignStatusPill(f.Status)}
        ${dates ? `<span>${escHtml(dates)}</span>` : ''}
        ${f.Region ? `<span>${escHtml(f.Region)}</span>` : ''}
        ${landing ? `<a class="cd-landing" href="${escAttr(landing)}" target="_blank" rel="noopener">View landing page →</a>` : ''}
      </div>
    </div>

    ${metrics.length ? `<div class="cd-metrics">${metrics.map(m => `
      <div class="cd-metric"><div class="cd-metric-label">${escHtml(m.label)}</div><div class="cd-metric-value">${m.value}</div></div>
    `).join('')}</div>` : ''}

    <p class="cd-blocks-title">Assets &amp; resources</p>
    <div class="cd-blocks" id="cd-blocks">
      <div class="skeleton-card" style="height:96px;border-radius:12px"></div>
      <div class="skeleton-card" style="height:96px;border-radius:12px"></div>
      <div class="skeleton-card" style="height:96px;border-radius:12px"></div>
    </div>

    <div id="cd-asset-panel" style="margin-top:20px"></div>`;

  // Remember what the asset blocks should resolve against, then discover
  // the item's real sub-folders (nothing hardcoded - whatever folders
  // marketing creates in SharePoint show up as blocks).
  _detailContext = { folderRoot: opts.folderRoot, campaignFolder: f.CampaignFolder || f.Folder || f.Title };
  _loadDetailBlocks();
}

const _CD_BLOCK_SVG = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="22" height="22"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>';

let _detailBlocks = [];

function _assetFilesHTML(heading, files, listVar) {
  return `
      <p class="cd-blocks-title" style="margin-bottom:10px">${escHtml(heading)}</p>
      <div class="asset-grid">${files.map((f, idx) => {
        const icon = fileIcon(f.name, false);
        const meta = [humanSize(f.size), fmtSpDate(f.lastModifiedDateTime)].filter(Boolean).join(' \u00b7 ');
        return `<div class="asset asset-preview" role="button" tabindex="0" onclick="openDocFile(${listVar}[${idx}])" onkeydown="if(event.key==='Enter')openDocFile(${listVar}[${idx}])">
          <div class="asset-icon ${icon.cls}">${escHtml(icon.label)}</div>
          <div class="asset-info"><div class="asset-name">${escHtml(f.name)}</div><div class="asset-meta">${escHtml(meta)}</div></div>
          <span class="asset-open">Open in hub \u2192</span>
        </div>`;
      }).join('')}</div>`;
}

async function _loadDetailBlocks() {
  const ctx  = _detailContext;
  const wrap = document.getElementById('cd-blocks');
  if (!ctx || !wrap) return;
  try {
    const drive = await resolveDrive(HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary);
    const rootFolder = await _findChildFolder(drive.id, null, ctx.folderRoot);
    if (!rootFolder) throw new Error(`No "${ctx.folderRoot}" folder in the document library yet.`);
    const itemFolder = await _findChildFolder(drive.id, rootFolder.id, ctx.campaignFolder);
    if (!itemFolder) throw new Error(`No folder matching "${ctx.campaignFolder}" inside ${ctx.folderRoot} yet - create one and its asset folders appear here automatically.`);
    if (_detailContext !== ctx) return; // user navigated away meanwhile

    const kids       = await fetchDriveChildren(drive.id, itemFolder.id);
    const folders    = kids.filter(k => k.folder);
    const looseFiles = kids.filter(k => !k.folder);

    _detailBlocks = folders.map(k => ({ label: k.name, id: k.id, driveId: drive.id, count: (k.folder && typeof k.folder.childCount === 'number') ? k.folder.childCount : null }));

    if (!_detailBlocks.length && !looseFiles.length) {
      wrap.innerHTML = '<p class="prose dim">No asset folders set up for this item yet - add folders inside its Documents folder and they appear here.</p>';
      return;
    }

    wrap.innerHTML = _detailBlocks.map((b, i) => `
      <div class="cd-block" role="button" tabindex="0" onclick="openDetailAsset(${i})" onkeydown="if(event.key==='Enter')openDetailAsset(${i})">
        <span class="cd-block-ico">${_CD_BLOCK_SVG}</span>
        ${escHtml(b.label)}
        <span class="cd-block-note">${b.count !== null ? b.count + (b.count === 1 ? ' item' : ' items') + ' \u00b7 ' : ''}Open in hub</span>
      </div>`).join('');

    if (looseFiles.length) {
      _looseDetailFiles = looseFiles;
      const panel = document.getElementById('cd-asset-panel');
      if (panel) panel.innerHTML = _assetFilesHTML('Other files', looseFiles, '_looseDetailFiles');
    }
  } catch (e) {
    wrap.innerHTML = `<p class="prose dim">${escHtml(e.message)}</p>`;
  }
}

let _looseDetailFiles = [];

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
function _nameTokens(s) {
  return String(s || '').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
}
async function _findChildFolder(driveId, parentId, name) {
  const target = String(name || '').trim().toLowerCase();
  if (!target) return null;
  const items = await fetchDriveChildren(driveId, parentId);
  const folders = items.filter(f => f.folder);
  // 1. exact (case-insensitive)
  let hit = folders.find(f => String(f.name).trim().toLowerCase() === target);
  if (hit) return hit;
  // 2. substring either way
  hit = folders.find(f => {
    const n = String(f.name).trim().toLowerCase();
    return n.includes(target) || target.includes(n);
  });
  if (hit) return hit;
  // 3. word-set match: same words in any order, or one a subset of the
  //    other. Fixes e.g. list title "Flat-Pack Tubular Black ... Stand range"
  //    vs folder "Black Flat-Pack Tubular ... Stand Range".
  const t = _nameTokens(target);
  if (!t.length) return null;
  return folders.find(f => {
    const n = _nameTokens(f.name);
    if (!n.length) return false;
    const setN = new Set(n), setT = new Set(t);
    return t.every(w => setN.has(w)) || n.every(w => setT.has(w));
  }) || null;
}

// Open a discovered asset block folder in-hub. One file opens straight
// into the preview; several are listed in a panel underneath the blocks.
async function openDetailAsset(i) {
  const block = _detailBlocks[i];
  const panel = document.getElementById('cd-asset-panel');
  if (!block || !panel) return;

  panel.innerHTML = `<p class="prose dim">Opening \u201c${escHtml(block.label)}\u201d\u2026</p>`;

  try {
    const kids  = await fetchDriveChildren(block.driveId, block.id);
    const files = kids.filter(x => !x.folder);
    const subs  = kids.filter(x => x.folder);

    if (!files.length && !subs.length) {
      panel.innerHTML = `<p class="prose dim">No files in \u201c${escHtml(block.label)}\u201d yet.</p>`;
      return;
    }

    // Sub-folders inside a block become extra blocks appended to the grid.
    if (subs.length) {
      const wrap = document.getElementById('cd-blocks');
      const baseIdx = _detailBlocks.length;
      subs.forEach((s, si) => {
        _detailBlocks.push({ label: block.label + ' / ' + s.name, id: s.id, driveId: block.driveId, count: (s.folder && typeof s.folder.childCount === 'number') ? s.folder.childCount : null });
        if (wrap) wrap.insertAdjacentHTML('beforeend', `
          <div class="cd-block" role="button" tabindex="0" onclick="openDetailAsset(${baseIdx + si})" onkeydown="if(event.key==='Enter')openDetailAsset(${baseIdx + si})">
            <span class="cd-block-ico">${_CD_BLOCK_SVG}</span>
            ${escHtml(block.label + ' / ' + s.name)}
            <span class="cd-block-note">Open in hub</span>
          </div>`);
      });
    }

    if (!files.length) {
      panel.innerHTML = `<p class="prose dim">\u201c${escHtml(block.label)}\u201d holds sub-folders \u2014 they've been added to the grid above.</p>`;
      return;
    }

    _lastAssetFiles = files;
    if (files.length === 1) {
      openDocFile(files[0]);
      panel.innerHTML = `<p class="prose dim">Opened <strong>${escHtml(files[0].name)}</strong> \u2014 <a class="fb-crumb" onclick="openDocFile(_lastAssetFiles[0])">reopen</a></p>`;
      return;
    }
    panel.innerHTML = _assetFilesHTML(`${block.label} \u2014 ${files.length} files`, files, '_lastAssetFiles');
  } catch (e) {
    panel.innerHTML = `<p class="prose dim">${escHtml(e.message)}</p>`;
  }
}

let _lastAssetFiles = [];
