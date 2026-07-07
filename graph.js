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

// Resolve the SharePoint site ID once per session
let _siteIdPromise = null;

function getSiteId() {
  if (_siteIdPromise) return _siteIdPromise;

  _siteIdPromise = (async () => {
    const cached = _cacheGet('siteId');
    if (cached) return cached;

    const u = new URL(HUB_CONFIG.sharepointSite);
    const data = await graphFetch(`/sites/${u.hostname}:${u.pathname}`);
    _cacheSet('siteId', data.id);
    return data.id;
  })();

  _siteIdPromise.catch(() => { _siteIdPromise = null; });
  return _siteIdPromise;
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
    `/drives/${drive.id}/root/children?$select=name,size,lastModifiedDateTime,webUrl,file,folder&$top=100`
  );
  const files = data.value || [];
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

function renderLaunches(items) {
  const el = document.getElementById('sp-launches-list');
  if (!el) return;

  if (!items.length) {
    el.innerHTML = '<p class="prose dim">No launches in SharePoint yet — add items to the Product Launches list.</p>';
    return;
  }

  const sorted = [...items].sort((a, b) => String(a.LaunchDate || '').localeCompare(String(b.LaunchDate || '')));

  el.innerHTML = `<div class="asset-grid">${sorted.map(f => {
    const url = safeUrl(linkOf(f.LinkURL), '');
    return `
    <div class="asset">
      <div class="asset-info">
        <div class="asset-name">${escHtml(f.Title || 'Untitled')}</div>
        <div class="asset-meta">${[
          escHtml(f.SKU || ''),
          fmtSpDate(f.LaunchDate),
          f.RRP != null && f.RRP !== '' ? fmtMoney(f.RRP) : '',
        ].filter(Boolean).join(' · ')}</div>
      </div>
      ${statusBadge(f.Status)}
      ${url ? `<a class="link" href="${escAttr(url)}" target="_blank" rel="noopener">Open →</a>` : ''}
    </div>`;
  }).join('')}</div>`;
}

function renderCampaigns(items) {
  const grid = document.getElementById('sp-campaigns-grid');
  if (!grid) return;

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
    const url = safeUrl(linkOf(f.LinkURL), '');

    return `
    <article class="camp-card" ${url ? `onclick="window.open('${escAttr(url)}','_blank','noopener')"` : ''}>
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

function renderDocuments(files) {
  const grid = document.getElementById('sp-documents-grid');
  if (!grid) return;

  if (!files.length) {
    grid.innerHTML = '<p class="prose dim">The library is empty — upload files to SharePoint and they\'ll appear here.</p>';
    return;
  }

  grid.innerHTML = `<div class="asset-grid">${files.map(f => {
    const icon = fileIcon(f.name, !!f.folder);
    const url = safeUrl(f.webUrl, '');
    const meta = [
      f.folder ? `${f.folder.childCount ?? ''} items`.trim() : humanSize(f.size),
      fmtSpDate(f.lastModifiedDateTime),
    ].filter(Boolean).join(' · ');

    return `
    <a class="asset" ${url ? `href="${escAttr(url)}" target="_blank" rel="noopener"` : ''}>
      <div class="asset-icon ${icon.cls}">${escHtml(icon.label)}</div>
      <div class="asset-info">
        <div class="asset-name">${escHtml(f.name)}</div>
        <div class="asset-meta">${escHtml(meta)}</div>
      </div>
    </a>`;
  }).join('')}</div>`;
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

  const vids = [
    ...(wp.status === 'fulfilled' ? wp.value : []),
    ...(sp.status === 'fulfilled' ? sp.value : []),
  ].sort((a, b) => String(b.date).localeCompare(String(a.date)))
   .slice(0, (HUB_CONFIG.videos && HUB_CONFIG.videos.max) || 6);

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

async function loadSharePointDocuments() {
  const grid = document.getElementById('sp-documents-grid');
  if (!grid) return;

  grid.innerHTML = '<div class="skeleton sk-line med"></div><div class="skeleton sk-line"></div><div class="skeleton sk-line short"></div>';

  try {
    const files = await fetchLibraryFiles();
    renderDocuments(files);
  } catch (e) {
    _renderListError('sp-documents-grid', `Couldn't load the document library: ${e.message}`);
  }
}
