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
// Every document library on a site, not just the one called
// "Documents". 1 Sep 2026: the Product Portal's six Product Change
// Notifications live in FormServerTemplates — a system library — because
// they were dropped onto a page instead of into the library. A crawl of
// "Documents" alone could never find them, which is exactly why the
// portal page looked like it had none.
async function resolveAllDrives(siteUrl) {
  const siteId = await resolveSiteId(siteUrl);
  const drives = await graphFetch(`/sites/${siteId}/drives?$select=id,name,webUrl`);
  return (drives.value || []);
}

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
  // 1 Sep 2026 — @microsoft.graph.downloadUrl is asked for HERE, with the
  // listing. It used to be fetched one file at a time when a picture came
  // into view, which is a Graph round trip per image: a campaign folder
  // with twenty pictures paid twenty of them before anything appeared.
  // Now the listing already has them and the page paints straight away.
  const data = await graphFetch(
    `${base}?$select=id,name,size,lastModifiedDateTime,webUrl,file,folder,@microsoft.graph.downloadUrl&$top=200`
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

// ═══ Imagery from SharePoint ═════════════════════════════════
// The launch and campaign cards used to be flat colour blocks, which
// is most of why the pages read as a list rather than a website. The
// artwork marketing need is already in SharePoint — in the item's own
// asset folder — so the hub goes and gets it.
//
// Graph's /thumbnails endpoint returns a PRE-AUTHENTICATED url: it
// works in a plain <img>/background-image with no token attached, and
// expires after a few hours, which is exactly right for a page that is
// re-rendered on every visit. Never cache these to disk.

const IMAGE_EXT = /\.(jpe?g|png|gif|webp|bmp|tiff?)$/i;
const _thumbCache = {};

async function driveThumb(driveId, itemId, size) {
  const want = size || 'large';
  const key  = driveId + '|' + itemId + '|' + want;
  if (_thumbCache[key] !== undefined) return _thumbCache[key];
  try {
    const d = await graphFetch(`/drives/${driveId}/items/${itemId}/thumbnails?$select=${want}`);
    const set = (d.value || [])[0] || {};
    const url = (set[want] && set[want].url) || '';
    _thumbCache[key] = url;
    return url;
  } catch (_) {
    _thumbCache[key] = '';
    return '';
  }
}

// Best image inside a folder. Prefers something obviously meant as the
// picture for the thing ("hero", "cover", "main", "banner"), then any
// image sitting loose in the folder, then the first image inside an
// images/artwork sub-folder.
async function folderHeroImage(driveId, folderId) {
  try {
    const kids = await fetchDriveChildren(driveId, folderId);

    // 1 Sep 2026 — David, relaying marketing: "images to be used for
    // each campaign/launch are saved as email campaign banner in the
    // email campaign folder for each launch/campaign". That folder is
    // therefore the answer, not a guess: it exists under both
    // Documents ▸ Campaigns ▸ <name> and Documents ▸ Launches ▸ <name>,
    // and the file is named "... Email Campaign Banner.png" (or .gif).
    // Checked first, so a stray photo loose in the campaign folder can
    // no longer win. Anything without one falls through to the old
    // behaviour untouched.
    const emailFolder = kids.find(k => k.folder && /e-?mail\s*camp/i.test(k.name || ''));
    if (emailFolder) {
      const inner = await fetchDriveChildren(driveId, emailFolder.id);
      const pics  = inner.filter(k => !k.folder && IMAGE_EXT.test(k.name || ''));
      if (pics.length) {
        const pick = pics.find(k => /banner|header/i.test(k.name)) || pics[0];
        const url  = await driveThumb(driveId, pick.id);
        if (url) return url;
      }
      console.info('[Imagery] "' + emailFolder.name + '" has no image in it yet.');
    }

    const imgs = kids.filter(k => !k.folder && IMAGE_EXT.test(k.name || ''));
    if (imgs.length) {
      const pick = imgs.find(k => /hero|cover|main|banner|key ?visual/i.test(k.name)) || imgs[0];
      return await driveThumb(driveId, pick.id);
    }
    const sub = kids.find(k => k.folder && /image|photo|artwork|visual|social|asset/i.test(k.name || ''));
    if (sub) {
      const inner = await fetchDriveChildren(driveId, sub.id);
      const first = inner.find(k => !k.folder && IMAGE_EXT.test(k.name || ''));
      if (first) return await driveThumb(driveId, first.id);
    }
  } catch (_) { /* no folder, no access — card keeps its fallback */ }
  return '';
}

// Names in SharePoint rarely match a list Title character for character.
// Compare on letters and digits only so "FX-90 Launch" finds "FX90".
function _slugKey(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

// Resolve a hero image for every item in a list, by finding its folder
// under Documents/<root>/. Returns a map keyed by the item Title.
// Everything is best-effort: one missing folder never stops the rest.
async function itemHeroImages(items, rootFolderName) {
  const out = {};
  if (!items || !items.length) return out;
  try {
    const drive = await resolveDrive(HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary);
    const root  = await _findChildFolder(drive.id, null, rootFolderName);
    if (!root) return out;

    const folders = (await fetchDriveChildren(drive.id, root.id)).filter(f => f.folder);
    const byKey = {};
    folders.forEach(f => { byKey[_slugKey(f.name)] = f; });

    await Promise.all(items.slice(0, 24).map(async f => {
      const wanted = f.CampaignFolder || f.Folder || f.Title;
      const key = _slugKey(wanted);
      if (!key) return;
      const folder = byKey[key]
        || folders.find(x => _slugKey(x.name).includes(key) || key.includes(_slugKey(x.name)));
      if (!folder) return;
      const url = await folderHeroImage(drive.id, folder.id);
      if (url) out[f.Title] = url;
    }));
  } catch (e) {
    console.info('[Imagery] no folder artwork available:', e.message);
  }
  return out;
}

// Paint the images in after the cards are already on screen, so the
// page never waits on Graph to show something.
function _applyHeroImages(prefix, items, map) {
  items.forEach((f, i) => {
    const url = map[f.Title];
    if (!url) return;
    const el = document.getElementById(prefix + i);
    if (!el) return;
    el.style.backgroundImage = `url('${safeCssUrl(url)}')`;
    el.classList.add('has-img');
  });
}

// ── Landing page artwork ──────────────────────────────────────
// WordPress *pages* almost never carry a featured image, so "Updated
// landing pages" on the home page had nothing to show. Marketing now
// drop artwork into Documents ▸ Images for Landing Pages instead, and
// the hub matches a picture to a page by filename. See config.js for
// how forgiving the matching is.
//
// Returns [{ key, url }]; app.js does the matching so the carousel can
// render immediately and take the pictures when they arrive.
let _landingImgs = null;

async function fetchLandingImages() {
  if (_landingImgs) return _landingImgs;
  const cfg = HUB_CONFIG.landingImages || {};
  if (!cfg.folder) return (_landingImgs = []);

  try {
    const site  = cfg.site === 'product' ? HUB_CONFIG.productPortalSite : HUB_CONFIG.sharepointSite;
    const drive = await resolveDrive(site, HUB_CONFIG.documentsLibrary);
    const folder = await _findChildFolder(drive.id, null, cfg.folder);
    if (!folder) {
      console.info(`[Landing images] no "${cfg.folder}" folder yet — cards stay text-only.`);
      return (_landingImgs = []);
    }

    // 26 Aug, second round: the pictures weren't showing because they
    // aren't loose in that folder — David: "I think you need to go 1
    // deeper then in another folder in images for landing pages". So
    // this walks SUB-FOLDERS too, and the sub-folder's own name counts
    // as a match key. A folder called "Fire Extinguishers" holding
    // "hero.jpg" now finds /fire-extinguishers, which is the shape
    // marketing were actually using.
    const found = [];
    const walk = async (itemId, trail, depth) => {
      if (depth < 0 || found.length > 200) return;
      const kids = await fetchDriveChildren(drive.id, itemId);
      const subs = [];
      for (const k of kids) {
        if (k.folder) { subs.push(k); continue; }
        if (!IMAGE_EXT.test(k.name || '')) continue;
        found.push({ item: k, trail: trail });
      }
      await Promise.all(subs.map(sf => walk(sf.id, trail.concat(sf.name), depth - 1)));
    };
    await walk(folder.id, [], (cfg.depth === undefined ? 3 : cfg.depth));

    if (!found.length) {
      console.info(`[Landing images] "${cfg.folder}" has no images in it yet.`);
      return (_landingImgs = []);
    }

    const rows = await Promise.all(found.map(async ({ item, trail }) => {
      const bare = String(item.name).replace(/\.[a-z0-9]+$/i, '');
      // Match on the filename, on the folder it sits in, and on the two
      // joined — so "Fire Extinguishers/hero.jpg", "fire-extinguishers.jpg"
      // and "Landing/Fire Extinguishers 01.png" all land on the same page.
      const keys = [];
      const push = v => { const k = _slugKey(v); if (k && keys.indexOf(k) < 0) keys.push(k); };
      push(bare);
      if (trail.length) {
        push(trail[trail.length - 1]);
        push(trail[trail.length - 1] + ' ' + bare);
      }
      // Words for the forgiving matcher — the filename and every folder
      // it sits inside both describe the page.
      const words = [_landingWords(bare)].concat(trail.map(_landingWords));
      return { keys, words, url: await driveThumb(drive.id, item.id), name: item.name, folder: trail.join('/') };
    }));

    _landingImgs = rows.filter(r => r.keys.length && r.url);
    console.info(`[Landing images] ${_landingImgs.length} image(s) available.`);
  } catch (e) {
    console.info('[Landing images] unavailable:', e.message);
    _landingImgs = [];
  }
  return _landingImgs;
}

// Pick the image for one page. Exact key first, then either name
// containing the other — so "01 fire-extinguishers hero.jpg" still
// finds /fire-extinguishers. The longest match wins, which stops
// "water" hijacking "water-mist".
//
// 1 Sep 2026 — "Image is still not pulling through" (Fire Equipment
// Suppliers). The picture was in SharePoint all along, in a folder
// called "Fire Equipment Supplier Landing Page". Squashed to a slug
// that is "fireequipmentsupplierlandingpage"; the page is
// "fireequipmentsuppliers". Neither string contains the other — one
// says "landingpage", the other has a plural "s" in the middle — so
// the old rule found nothing and the card stayed blank. Sister page
// "Fire Extinguisher Supplier" matched only because its folder name
// happened to line up.
//
// Comparing WORDS fixes the whole class of problem: drop the words that
// say nothing about which page this is ("landing", "page", "banner"…),
// ignore a trailing "s", and score on how many real words the two share.
function _landingWords(s) {
  const noise = (HUB_CONFIG.landingImages && HUB_CONFIG.landingImages.noiseWords) || [];
  const stop  = noise.map(w => String(w).toLowerCase());
  return String(s || '')
    .toLowerCase()
    .replace(/\.[a-z0-9]+$/i, '')
    .split(/[^a-z0-9]+/)
    .filter(Boolean)
    .map(w => w.replace(/s$/, ''))            // supplier / suppliers
    .filter(w => w.length > 1 && stop.indexOf(w) < 0 && stop.indexOf(w + 's') < 0);
}

function matchLandingImage(images, page) {
  if (!images || !images.length) return '';
  let slug = '';
  try {
    const p = new URL(page.link).pathname.replace(/\/+$/, '');
    slug = p.split('/').filter(Boolean).pop() || '';
  } catch (_) { /* fall through to the title */ }
  const wanted = [_slugKey(slug), _slugKey(page.title)].filter(Boolean);

  // 1. The old exact-slug hit still wins outright when there is one.
  for (const w of wanted) {
    const exact = images.find(im => im.keys.indexOf(w) >= 0);
    if (exact) return exact.url;
  }

  // 2. Otherwise score on shared words.
  const pageWords = [...new Set(_landingWords(slug).concat(_landingWords(page.title)))];
  if (pageWords.length) {
    const need = (HUB_CONFIG.landingImages && HUB_CONFIG.landingImages.minWordMatch) || 1;
    let best = null, bestScore = 0;
    for (const im of images) {
      const imWords = [...new Set([].concat(...(im.words || [])))];
      if (!imWords.length) continue;
      const shared = pageWords.filter(w => imWords.indexOf(w) >= 0).length;
      if (!shared) continue;
      // Ratio breaks ties: two words out of two beats two out of nine.
      const score = shared + shared / Math.max(imWords.length, pageWords.length);
      if (shared >= need && score > bestScore) { best = im; bestScore = score; }
    }
    if (best) return best.url;
  }

  // 3. Last resort — the original containment rule, so nothing that
  //    used to match stops matching.
  let best = null, bestLen = 0;
  for (const w of wanted) {
    for (const im of images) {
      for (const k of im.keys) {
        if ((k.includes(w) || w.includes(k)) && k.length > bestLen) { best = im; bestLen = k.length; }
      }
    }
  }
  return best ? best.url : '';
}

// ═══ Renderers ═══════════════════════════════════════════════

// ── Product launches & campaigns: the editorial layout ────────
//
// REWRITTEN 26 Aug 2026. These two pages were a metric rail above
// three columns of small cards, which is a project tracker, not a
// website — marketing's words were "it doesn't look right". They now
// share one editorial shape:
//
//   1. a lead spread — the launch/campaign that matters right now,
//      given the space a homepage feature would get
//   2. a filter rail carrying the same traffic-light counts the metric
//      rail used to, but as something you can actually click
//   3. a generous card grid with real artwork out of SharePoint
//
// The RAG colours are unchanged, so red still means planning wherever
// you see it, and openLaunchDetail(i) / openCampaignDetail(i) still
// take the index into the sorted array.

// Cards carry a status dot rather than a filled block, so the artwork
// underneath is the thing you notice first.
function _pxDot(status) {
  return `<span class="px-badge"><span class="px-badge-dot ${ragOf(status)}"></span>${escHtml(status || 'Not set')}</span>`;
}

// A card with no artwork isn't left grey — it gets the item's own
// initials set large on the brand gradient, which reads as a designed
// placeholder rather than a missing image.
function _pxInitials(title) {
  // Words that START with a letter only — "FX-90 Water Mist" should
  // read FW, not F9.
  const words = String(title || '').replace(/[^A-Za-z0-9 ]/g, ' ')
    .split(/\s+/).filter(w => /^[A-Za-z]/.test(w));
  return words.slice(0, 2).map(w => w[0]).join('').toUpperCase() || 'CF';
}

function _pxCard(kind, f, i, opts) {
  const o = opts || {};
  const fn = kind === 'launch' ? 'openLaunchDetail' : 'openCampaignDetail';
  return `
    <article class="px-card" data-tone="${ragOf(f.Status)}" style="--i:${i}"
             role="button" tabindex="0"
             onclick="${fn}(${i})" onkeydown="if(event.key==='Enter')${fn}(${i})">
      <div class="px-card-media" id="px-img-${kind}-${i}">
        <span class="px-card-initials">${escHtml(_pxInitials(f.Title))}</span>
        ${_pxDot(f.Status)}
      </div>
      <div class="px-card-body">
        <div class="px-card-eyebrow">${escHtml(o.eyebrow || 'Product launch')}</div>
        <h3 class="px-card-title">${escHtml(f.Title || 'Untitled')}</h3>
        <div class="px-card-meta">${escHtml(o.meta || '')}</div>
        ${o.footer || ''}
      </div>
      <span class="px-card-go">Open <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="13" height="13"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg></span>
    </article>`;
}

// The filter rail. Counts come from live data and a chip with nothing
// behind it is disabled rather than hidden, so the shape of the page
// doesn't jump about as statuses change.
function _pxRail(kind, items, tones, note) {
  const count = t => items.filter(f => ragOf(f.Status) === t).length;
  const chips = [{ tone: 'all', label: 'All' }].concat(tones).map(c => {
    const n = c.tone === 'all' ? items.length : count(c.tone);
    return `<button class="px-chip${c.tone === 'all' ? ' active' : ''}${n ? '' : ' empty'}"
              data-tone="${escAttr(c.tone)}" ${n ? '' : 'disabled'}
              onclick="filterPx('${kind}','${escAttr(c.tone)}',this)">
              ${c.tone === 'all' ? '' : `<span class="px-chip-dot ${escAttr(c.tone)}"></span>`}
              ${escHtml(c.label)}<b>${n}</b>
            </button>`;
  }).join('');
  return `<div class="px-rail">
    <div class="px-chips">${chips}</div>
    ${note ? `<div class="px-rail-note">${note}</div>` : ''}
  </div>`;
}

function filterPx(kind, tone, btn) {
  const grid = document.getElementById(kind === 'launch' ? 'px-launch-grid' : 'px-camp-grid');
  if (grid) grid.setAttribute('data-filter', tone);
  if (btn && btn.parentElement) {
    btn.parentElement.querySelectorAll('.px-chip').forEach(b => b.classList.toggle('active', b === btn));
  }
}

// The lead spread. Big type, big picture, one clear action.
function _pxLead(kind, f, i, opts) {
  const o = opts || {};
  const fn = kind === 'launch' ? 'openLaunchDetail' : 'openCampaignDetail';
  const codes = kind === 'launch' ? productCodes(f) : [];
  return `
  <section class="px-lead">
    <div class="px-lead-copy">
      <div class="px-eyebrow">${escHtml(o.eyebrow || 'Up next')}</div>
      <h2 class="px-lead-title">${escHtml(f.Title || 'Untitled')}</h2>
      ${o.sub ? `<p class="px-lead-sub">${escHtml(o.sub)}</p>` : ''}
      ${codes.length ? `<div class="px-lead-codes">${codes.map(c => `<span class="px-code">${escHtml(c)}</span>`).join('')}</div>` : ''}
      <div class="px-lead-meta">
        ${_pxDot(f.Status)}
        ${o.meta ? `<span class="px-lead-when">${escHtml(o.meta)}</span>` : ''}
      </div>
      <button class="px-cta" onclick="${fn}(${i})">
        ${escHtml(o.cta || 'View the launch')}
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
      </button>
    </div>
    <div class="px-lead-media" id="px-lead-${kind}">
      <span class="px-lead-initials">${escHtml(_pxInitials(f.Title))}</span>
    </div>
  </section>`;
}

let _launchItems = [];

function renderLaunches(items) {
  const el = document.getElementById('sp-launches-list');
  if (!el) return;

  items = items || [];

  if (!items.length) {
    el.innerHTML = `<div class="px-empty">
      <h3>No launches yet</h3>
      <p>Add items to the <strong>Product Launches</strong> list on the MarketingHub SharePoint site and they appear here — with the artwork from their asset folder.</p>
    </div>`;
    return;
  }

  // Newest first so the current launch leads the page.
  const sorted = [...items].sort((a, b) =>
    String(b.LaunchDate || '').localeCompare(String(a.LaunchDate || '')));
  _launchItems = sorted;

  // The lead is the next launch still to come; if they're all behind
  // us, the most recent one.
  const now = new Date();
  let leadIdx = sorted.findIndex(f => f.LaunchDate && new Date(f.LaunchDate) >= now);
  if (leadIdx < 0) leadIdx = 0;
  // Among future launches, feature the SOONEST rather than the furthest.
  const future = sorted
    .map((f, i) => ({ f, i }))
    .filter(x => x.f.LaunchDate && new Date(x.f.LaunchDate) >= now)
    .sort((a, b) => String(a.f.LaunchDate).localeCompare(String(b.f.LaunchDate)));
  if (future.length) leadIdx = future[0].i;

  const lead = sorted[leadIdx];
  const isFuture = lead.LaunchDate && new Date(lead.LaunchDate) >= now;

  el.innerHTML =
    _pxLead('launch', lead, leadIdx, {
      eyebrow: isFuture ? 'Next launch' : 'Most recent launch',
      sub: lead.Description || lead.Summary || '',
      meta: fmtSpDate(lead.LaunchDate) || 'Date to be confirmed',
      cta: 'View the launch',
    }) +
    _pxRail('launch', sorted,
      [{ tone: 'red', label: 'Planning' },
       { tone: 'amber', label: 'Confirmed' },
       { tone: 'green', label: 'Launched' },
       { tone: 'grey', label: 'Archive' }]) +
    `<div class="px-grid" id="px-launch-grid" data-filter="all">${
      sorted.map((f, i) => {
        const codes = productCodes(f);
        return _pxCard('launch', f, i, {
          eyebrow: 'Product launch',
          meta: fmtSpDate(f.LaunchDate) || 'Date to be confirmed',
          footer: codes.length
            ? `<div class="px-card-codes">${codes.slice(0, 4).map(c => `<span class="px-code sm">${escHtml(c)}</span>`).join('')}</div>`
            : '',
        });
      }).join('')
    }</div>`;

  // Artwork arrives after the page is already usable.
  itemHeroImages(sorted, (HUB_CONFIG.folders && HUB_CONFIG.folders.launches) || 'Launches')
    .then(map => {
      _applyHeroImages('px-img-launch-', sorted, map);
      const leadEl = document.getElementById('px-lead-launch');
      if (leadEl && map[lead.Title]) {
        leadEl.style.backgroundImage = `url('${safeCssUrl(map[lead.Title])}')`;
        leadEl.classList.add('has-img');
      }
    });
}

let _campaignItems = [];

function renderCampaigns(items) {
  const grid = document.getElementById('sp-campaigns-grid');
  if (!grid) return;

  items = items || [];
  _campaignItems = items;

  if (!items.length) {
    grid.classList.remove('camp-grid');
    grid.innerHTML = `<div class="px-empty">
      <h3>No campaigns yet</h3>
      <p>Add items to the <strong>Campaigns</strong> list on the MarketingHub SharePoint site. Drop artwork into <strong>Documents ▸ Campaigns ▸ &lt;campaign name&gt;</strong> and it becomes the card image.</p>
    </div>`;
    return;
  }

  // The old flat layout left .camp-grid (itself a 3-column grid) on this
  // element; writing a grid into a grid cell squashed every card to a
  // ninth of the page. Stripped in JS, not the HTML, so the loading
  // skeleton still lays out before this runs.
  grid.classList.remove('camp-grid');

  // Newest first by start date.
  const sorted = [...items].sort((a, b) =>
    String(b.StartDate || '').localeCompare(String(a.StartDate || '')));
  _campaignItems = sorted;

  // Lead with something live if there is one.
  let leadIdx = sorted.findIndex(f => ragOf(f.Status) === 'green');
  if (leadIdx < 0) leadIdx = 0;
  const lead = sorted[leadIdx];

  // No KPI strip. Emails sent / social posts / blogs / PR activity are
  // not tracked anywhere and there is no plan to track them, so they
  // only ever rendered as zeroes. Removed on David's instruction,
  // along with the metrics band on the detail page.

  grid.innerHTML =
    _pxLead('campaign', lead, leadIdx, {
      eyebrow: ragOf(lead.Status) === 'green' ? 'Running now' : 'Latest campaign',
      sub: lead.Description || lead.Summary || lead.CampaignType || '',
      meta: [fmtSpDate(lead.StartDate), fmtSpDate(lead.EndDate)].filter(Boolean).join(' – ')
            + (lead.Region ? ' · ' + lead.Region : ''),
      cta: 'Open the campaign',
    }) +
    _pxRail('campaign', sorted,
      [{ tone: 'red', label: 'Planning' },
       { tone: 'amber', label: 'Scheduled' },
       { tone: 'green', label: 'Live' },
       { tone: 'grey', label: 'Completed' }]) +
    `<div class="px-grid" id="px-camp-grid" data-filter="all">${
      sorted.map((f, i) => {
        const channels = Array.isArray(f.Channels) ? f.Channels : String(f.Channels || '').split(/[,;/]+/);
        const chips = channels.map(c => String(c).trim()).filter(Boolean).slice(0, 4);
        return _pxCard('campaign', f, i, {
          eyebrow: f.CampaignType || 'Campaign',
          meta: [fmtSpDate(f.StartDate), fmtSpDate(f.EndDate)].filter(Boolean).join(' – ')
                + (f.Region ? ' · ' + f.Region : ''),
          footer: chips.length
            ? `<div class="px-card-codes">${chips.map(c => `<span class="px-chan">${escHtml(c)}</span>`).join('')}</div>`
            : (f.Budget ? `<div class="px-card-codes"><span class="px-chan">${fmtMoney(f.Budget)}</span></div>` : ''),
        });
      }).join('')
    }</div>`;

  itemHeroImages(sorted, (HUB_CONFIG.folders && HUB_CONFIG.folders.campaigns) || 'Campaigns')
    .then(map => {
      _applyHeroImages('px-img-campaign-', sorted, map);
      const leadEl = document.getElementById('px-lead-campaign');
      if (leadEl && map[lead.Title]) {
        leadEl.style.backgroundImage = `url('${safeCssUrl(map[lead.Title])}')`;
        leadEl.classList.add('has-img');
      }
    });
}

// ═══ The reader ══════════════════════════════════════════════
//
// REWRITTEN 26 Aug 2026 (second round). David: "I still don't like how
// things open from SharePoint — is there a way this can open like a
// website would open?"
//
// It used to embed SharePoint's own preview iframe, which brings
// SharePoint's toolbars, branding and behaviour along with it. Now the
// hub fetches the file itself and renders it:
//
//   PDF            → the browser's own PDF viewer, off a blob: URL
//   image          → an <img>
//   text/csv/md    → set as text
//   video / audio  → the browser's own player
//   Word / PowerPoint / Excel
//                  → Graph is asked to CONVERT to PDF, then the same
//                    native viewer. Only if that fails does it fall
//                    back to the Office web viewer.
//
// Because the file becomes a blob: URL on our own origin there is no
// SharePoint chrome anywhere, and Download / Copy link / Open in new
// tab all work off it. Blob URLs are revoked when the reader closes.

const RDR = { blobUrl: null, file: null };

const RDR_PDF    = /\.pdf$/i;
const RDR_IMG    = /\.(png|jpe?g|gif|webp|bmp|svg|avif)$/i;
const RDR_TXT    = /\.(txt|md|csv|tsv|json|log|xml|ya?ml)$/i;
const RDR_AV     = /\.(mp4|webm|m4v|mov|mp3|m4a|wav|ogg)$/i;
const RDR_OFFICE = /\.(docx?|pptx?|xlsx?|xlsm|rtf|odt|odp|ods)$/i;

// 60 MB. Past that a blob is a bad idea in a browser tab — the reader
// hands the user the file instead of trying to paint it.
const RDR_MAX_BYTES = 60 * 1024 * 1024;

function _rdrExt(name) {
  return (String(name || '').split('.').pop() || '').toLowerCase();
}

// Graph's downloadUrl is PRE-AUTHENTICATED and short-lived: fetching it
// needs no Authorization header, which also means no CORS preflight.
async function _rdrDownloadUrl(f) {
  // Already have it from the folder listing? Use it. (These expire within
  // the hour, which is far longer than anyone keeps a page open.)
  if (f && f['@microsoft.graph.downloadUrl']) return f['@microsoft.graph.downloadUrl'];
  const meta = await graphFetch(
    `/drives/${f._driveId}/items/${f.id}?$select=id,name,size,@microsoft.graph.downloadUrl`);
  return (meta && meta['@microsoft.graph.downloadUrl']) || '';
}

async function _rdrBlob(f, asPdf) {
  if (asPdf) {
    // Graph converts Office formats to PDF on the fly. It answers with
    // a redirect to a pre-authenticated URL, and the Authorization
    // header is dropped on that cross-origin hop — which is fine,
    // because the target doesn't want it.
    const token = await getAccessToken();
    const res = await fetch(
      `${GRAPH_BASE}/drives/${f._driveId}/items/${f.id}/content?format=pdf`,
      { headers: { Authorization: 'Bearer ' + token } });
    if (!res.ok) throw new Error('Conversion returned ' + res.status);
    return res.blob();
  }
  const url = await _rdrDownloadUrl(f);
  if (!url) throw new Error('No download URL');
  const res = await fetch(url);
  if (!res.ok) throw new Error('Download returned ' + res.status);
  return res.blob();
}

function _rdrRevoke() {
  if (RDR.blobUrl) { try { URL.revokeObjectURL(RDR.blobUrl); } catch (_) {} }
  RDR.blobUrl = null;
}

function _rdrStage(html) {
  const stage = document.getElementById('doc-stage');
  if (stage) stage.innerHTML = html;
}

// Where Back goes, and what it's called. Captured the moment a file is
// opened so the reader can hand you back to the exact page — and scroll
// position — you left, which is what makes it feel like a website
// rather than something that popped up at you.
let _rdrReturn = { id: 'home', scroll: 0 };

const _RDR_BACK = {
  home:      'Back to the home page',
  launches:  'Back to product launches',
  campaigns: 'Back to campaigns',
  trade:     'Back to trade & events',
  training:  'Back to resources',
  portal:    'Back to the product portal',
  search:    'Back to the search results',
};

function _rdrGoToPage() {
  const active = document.querySelector('.page.active');
  const from   = active ? active.id.replace(/^page-/, '') : 'home';

  // Opening a second file from inside the reader must not make Back
  // point at the reader itself.
  if (from !== 'reader') {
    _rdrReturn = { id: from, scroll: window.scrollY || window.pageYOffset || 0 };
  }

  const lbl = document.getElementById('rdr-back-label');
  if (lbl) lbl.textContent = _RDR_BACK[_rdrReturn.id] || 'Back';

  // Ember's panel would otherwise sit over the document she found.
  if (typeof toggleEmber === 'function') toggleEmber(false);

  if (typeof showPage === 'function') showPage('reader');
  // showPage clears the nav; keep the section you came from lit, because
  // as far as the reader is concerned you are still inside it.
  if (typeof updateNavActive === 'function') updateNavActive(_rdrReturn.id);
}

async function openDocFile(f) {
  if (!f) return;
  const page = document.getElementById('page-reader');
  if (!page) { if (f.webUrl) window.open(safeUrl(f.webUrl, '#'), '_blank', 'noopener'); return; }

  _rdrRevoke();
  RDR.file = f;

  const ext = _rdrExt(f.name);
  const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  set('rdr-name', f.name || 'Document');
  set('rdr-type', (ext || 'file').toUpperCase().slice(0, 4));
  set('rdr-sub', [humanSize(f.size), fmtSpDate(f.lastModifiedDateTime)].filter(Boolean).join(' · '));

  _rdrGoToPage();
  _rdrStage('<div class="rdr-wait"><span class="rdr-spin"></span>Opening…</div>');

  if (!f._driveId || !f.id) {
    _rdrStage(`<div class="rdr-wait">This one can only be opened in SharePoint.
      <a class="rdr-alt" href="${escAttr(safeUrl(f.webUrl, '#'))}" target="_blank" rel="noopener">Open it there →</a></div>`);
    return;
  }

  // ── The plan, in order of preference ──────────────────────
  // Images, video and audio never need the file's BYTES — an <img> or
  // <video> src doesn't do CORS, so the pre-authenticated downloadUrl
  // goes straight into the tag and just works.
  //
  // PDFs and text DO need the bytes (blob → the browser's own PDF
  // viewer — the "opens like a website" experience). But fetching the
  // downloadUrl is a cross-origin fetch, and whether SharePoint sends
  // CORS headers on it varies by tenant — which is exactly why David
  // saw "This file type can't be shown" on a plain 445 KB PDF. So:
  //
  //   1. try the blob            (best: native viewer, no chrome)
  //   2. fall back to the Graph  (always works — it's an iframe, and
  //      preview iframe, nb=true  it's what the hub used for months)
  //   3. only then offer download
  //
  // Every failure is logged with the REAL reason, so the console says
  // what actually happened instead of a generic shrug.

  try {
    if (Number(f.size) > RDR_MAX_BYTES) {
      _rdrFallbackPreview(f, 'the file is over the ' + Math.round(RDR_MAX_BYTES / 1048576) + ' MB in-page limit');
      return;
    }

    if (RDR_IMG.test(f.name)) {
      const url = await _rdrDownloadUrl(f);
      if (!url) throw new Error('no download URL');
      // If even the direct URL won't render (rare), fall back to the
      // large thumbnail, which is served from a different, always-
      // embeddable endpoint.
      _rdrStage(`<div class="rdr-centre"><img class="rdr-img" src="${escAttr(url)}" alt="${escAttr(f.name)}"
        onerror="_rdrImgFallback(this)"></div>`);

    } else if (RDR_AV.test(f.name)) {
      const url = await _rdrDownloadUrl(f);
      if (!url) throw new Error('no download URL');
      const audio = /\.(mp3|m4a|wav|ogg)$/i.test(f.name);
      _rdrStage(audio
        ? `<div class="rdr-centre"><audio class="rdr-audio" src="${escAttr(url)}" controls></audio></div>`
        : `<div class="rdr-centre"><video class="rdr-video" src="${escAttr(url)}" controls playsinline></video></div>`);

    } else if (RDR_PDF.test(f.name)) {
      try {
        RDR.blobUrl = URL.createObjectURL(await _rdrBlob(f, false));
        _rdrStage(`<iframe class="rdr-frame" src="${escAttr(RDR.blobUrl)}#view=FitH" title="${escAttr(f.name)}"></iframe>`);
      } catch (e) {
        console.warn('[Reader] PDF blob route failed (' + e.message + ') — using the preview service.');
        await _rdrFallbackPreview(f);
      }

    } else if (RDR_TXT.test(f.name)) {
      try {
        const text = await (await _rdrBlob(f, false)).text();
        _rdrStage(`<pre class="rdr-text">${escHtml(text.slice(0, 400000))}</pre>`);
      } catch (e) {
        console.warn('[Reader] text blob route failed (' + e.message + ') — using the preview service.');
        await _rdrFallbackPreview(f);
      }

    } else if (RDR_OFFICE.test(f.name)) {
      try {
        RDR.blobUrl = URL.createObjectURL(await _rdrBlob(f, true));
        _rdrStage(`<iframe class="rdr-frame" src="${escAttr(RDR.blobUrl)}#view=FitH" title="${escAttr(f.name)}"></iframe>`);
      } catch (e) {
        console.warn('[Reader] Office→PDF conversion failed (' + e.message + ') — using the preview service.');
        await _rdrFallbackPreview(f);
      }

    } else {
      // A type nothing can paint (zip, exe, font…) — go straight to
      // the honest offer.
      _rdrStage(`<div class="rdr-wait">This file type can’t be shown in the page.
        <button class="rdr-alt" onclick="downloadDoc()">Download it instead ↓</button></div>`);
    }
  } catch (e) {
    console.warn('[Reader] could not open ' + f.name + ':', e.message);
    await _rdrFallbackPreview(f);
  }
}

// The safety net: Microsoft's preview service in an iframe. Renders
// essentially everything Office knows about, needs no CORS (iframes
// don't), and nb=true strips most of the branding. Not as clean as the
// native viewer, but a document ALWAYS beats an apology.
async function _rdrFallbackPreview(f, why) {
  try {
    const prev = await graphPost(`/drives/${f._driveId}/items/${f.id}/preview`, {});
    const url = prev && prev.getUrl;
    if (!url) throw new Error('no preview URL');
    _rdrStage(`<iframe class="rdr-frame" src="${escAttr(url + (url.includes('?') ? '&' : '?') + 'nb=true')}" title="${escAttr(f.name)}"></iframe>`);
  } catch (e) {
    console.warn('[Reader] preview service also failed:', e.message);
    _rdrStage(`<div class="rdr-wait">${escHtml(why ? 'This can’t be shown in the page — ' + why + '.' : 'This document can’t be shown in the page right now.')}
      <button class="rdr-alt" onclick="downloadDoc()">Download it instead ↓</button></div>`);
  }
}

// An image whose direct URL refused to render — swap in the large
// thumbnail, which comes from an endpoint built for embedding.
async function _rdrImgFallback(imgEl) {
  const f = RDR.file;
  if (!f) return;
  try {
    const url = await driveThumb(f._driveId, f.id, 'large');
    if (url && imgEl) { imgEl.onerror = null; imgEl.src = url; return; }
  } catch (_) {}
  _rdrFallbackPreview(f);
}

// Download works off the blob when we already have it (instant, no
// second round trip) and off a fresh pre-authenticated URL otherwise.
async function downloadDoc() {
  const f = RDR.file;
  if (!f) return;
  try {
    const href = RDR.blobUrl || await _rdrDownloadUrl(f);
    if (!href) throw new Error('no url');
    const a = document.createElement('a');
    a.href = href;
    a.download = f.name || 'document';
    document.body.appendChild(a);
    a.click();
    a.remove();
  } catch (_) {
    if (f.webUrl) window.open(safeUrl(f.webUrl, '#'), '_blank', 'noopener');
  }
}

// "Share" copies the SharePoint link, not the blob — a blob: URL only
// exists inside this tab, so sending someone one would be useless. The
// SharePoint link opens for anyone who already has access to the site.
async function shareDoc(btn) {
  const f = RDR.file;
  if (!f || !f.webUrl) return;
  const url = safeUrl(f.webUrl, '');
  try {
    await navigator.clipboard.writeText(url);
    if (btn) {
      const t = btn.getAttribute('data-label') || btn.textContent;
      btn.setAttribute('data-label', t);
      btn.textContent = 'Link copied';
      setTimeout(() => { btn.textContent = t; }, 1800);
    } else if (typeof showToast === 'function') {
      showToast('Link copied');
    }
  } catch (_) {
    window.prompt('Copy this link:', url);
  }
}

// Opens the file itself in a new browser tab — a plain PDF in the
// browser's own viewer, with neither the hub nor SharePoint around it.
function popOutDoc() {
  if (RDR.blobUrl) { window.open(RDR.blobUrl, '_blank', 'noopener'); return; }
  const f = RDR.file;
  if (f && f.webUrl) window.open(safeUrl(f.webUrl, '#'), '_blank', 'noopener');
}

// "Back" from the reader. Not a close button — a navigation, which is
// why it restores the scroll position too: you land back on the row you
// clicked, not at the top of the page.
function closeDocPreview() {
  const page = document.getElementById('page-reader');
  if (!page) return;

  _rdrStage('');
  _rdrRevoke();
  RDR.file = null;

  const back = _rdrReturn || { id: 'home', scroll: 0 };
  if (typeof showPage === 'function') showPage(back.id);
  if (typeof updateNavActive === 'function') updateNavActive(back.id);
  // showPage scrolls to the top on the way in, so put us back after it.
  if (back.scroll) setTimeout(() => window.scrollTo({ top: back.scroll, behavior: 'auto' }), 140);
}

function _readerIsOpen() {
  const page = document.getElementById('page-reader');
  return !!(page && page.classList.contains('active'));
}

// Escape goes back, the same as the Back link.
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  if (_readerIsOpen()) closeDocPreview();
});

// ── Download / share on every file row ────────────────────────
// Every list of files in the hub — Resources, the Product Portal,
// campaign assets, event packs — gets the same two buttons, so staff
// have everything in one place without opening the file first.
// Files live in a small registry so a button only has to carry a key.
const DOCREG = {};
let _docRegSeq = 0;

function regDoc(f) {
  const k = 'd' + (++_docRegSeq);
  DOCREG[k] = f;
  return k;
}

function openRegDoc(k) { return openDocFile(DOCREG[k]); }

async function downloadRegDoc(k, ev) {
  if (ev && ev.stopPropagation) ev.stopPropagation();
  const f = DOCREG[k];
  if (!f) return;
  const keepFile = RDR.file, keepUrl = RDR.blobUrl;
  RDR.file = f; RDR.blobUrl = null;      // force a fresh pre-auth URL
  await downloadDoc();
  RDR.file = keepFile; RDR.blobUrl = keepUrl;
}

async function shareRegDoc(k, ev) {
  if (ev && ev.stopPropagation) ev.stopPropagation();
  const f = DOCREG[k];
  if (!f || !f.webUrl) return;
  const keepFile = RDR.file;
  RDR.file = f;
  await shareDoc(null);
  RDR.file = keepFile;
}

function docActions(k) {
  return `<span class="doc-acts">
    <button class="doc-act" title="Download" aria-label="Download" onclick="downloadRegDoc('${k}',event)">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
    </button>
    <button class="doc-act" title="Copy link" aria-label="Copy link" onclick="shareRegDoc('${k}',event)">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>
    </button>
    <button class="doc-act" title="Open full screen" aria-label="Open full screen" onclick="event.stopPropagation();openRegDoc('${k}')">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>
    </button>
  </span>`;
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
// Two possible sources, tried in order:
//   1. the new checkfire-ai Function app — reads the channel's public
//      RSS feed, so no API key and no Google account is involved
//   2. the old checkfire-jotform /api/videos endpoint, in case its
//      YOUTUBE_API_KEY ever gets set
// Whichever answers first wins. Nothing is embedded any more — the web
// filter blocks the in-page player — so all we need from either is a
// title, a date, a thumbnail and a link out to YouTube.
async function fetchYouTubeVideos() {
  const cfg = (HUB_CONFIG.videos && HUB_CONFIG.videos.youtube) || {};
  const ai  = ((HUB_CONFIG.ember || {}).aiProxyUrl || '').replace(/\/+$/, '');

  const urls = [];
  if (ai) urls.push(ai + '/videos');
  if (cfg.proxyUrl) urls.push(cfg.proxyUrl);
  if (!urls.length) return [];

  const cached = _cacheGet('videos_yt');
  if (cached) return cached;

  let rows = null;
  for (const url of urls) {
    try {
      const res = await fetch(url, { headers: { Accept: 'application/json' } });
      if (!res.ok) continue;
      const data = await res.json();
      const list = Array.isArray(data) ? data : (data.videos || data.items || []);
      if (list && list.length) { rows = list; break; }
    } catch (_) { /* try the next one */ }
  }
  if (!rows) return [];

  const vids = rows.map(v => {
    const id = v.id || v.videoId || '';
    return {
      title:     String(v.title || 'Untitled').trim(),
      date:      v.date || v.published || v.publishedAt || '',
      youtubeId: id,
      src:       null,
      href:      v.url || v.link || (id ? 'https://www.youtube.com/watch?v=' + encodeURIComponent(id) : ''),
      // i.ytimg.com serves thumbnails without a key. If the feed didn't
      // give us one, build it from the video id.
      thumb:     v.thumb || v.thumbnail ||
                 (id ? `https://i.ytimg.com/vi/${encodeURIComponent(id)}/mqdefault.jpg` : ''),
      source:    'YouTube',
    };
  }).filter(v => v.youtubeId);

  _cacheSet('videos_yt', vids);
  return vids;
}

async function loadHomeVideos() {
  // The hero box is the only place videos appear now — the grid and
  // the embedded player are both gone.
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

  // Prefer recent videos (default: last 3 months — see config.js), but
  // never at the cost of showing nothing: the box's whole job is "the
  // most recent upload", so if the age filter empties the list we keep
  // the unfiltered set and just show the newest.
  const months = (HUB_CONFIG.videos && HUB_CONFIG.videos.maxAgeMonths) || 0;
  if (months > 0) {
    const cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - months);
    const recent = vids.filter(v => v.date && !isNaN(new Date(v.date)) && new Date(v.date) >= cutoff);
    if (recent.length) vids = recent;
  }
  vids = vids.slice(0, (HUB_CONFIG.videos && HUB_CONFIG.videos.max) || 6);

  // The hero box, which is the only place videos appear now.
  renderHeroVideos(vids);
}

// ── Latest videos (home hero box) ─────────────────────────────
//
// REWRITTEN 26 Aug 2026 (round 2, then fixed). Two things were wrong
// on the live site:
//
//  1. The card was a full-width 16:9 thumbnail. This box only gets
//     about 200px of the hero grid, and styles.css pins that grid to
//     height:420px — so the card overflowed and painted over the row
//     underneath. That was the "sizing has gone off" break. The card
//     is horizontal now and fits.
//
//  2. It was showing a Media Portal video, so clicking "Latest"
//     dropped you into SharePoint. This box is YOUTUBE ONLY now —
//     anything else is filtered out here as well as switched off in
//     config, so it can't come back by accident. With no YouTube feed
//     it says so and offers the channel, rather than substituting
//     something that isn't a YouTube video.
function renderHeroVideos(vids) {
  const el = document.getElementById('home-hero-videos-body');
  if (!el) return;

  const yt = (HUB_CONFIG.videos && HUB_CONFIG.videos.youtube) || {};
  const channel = safeUrl(yt.channelUrl || '', '');

  const channelLink = channel
    ? `<a class="hbox-more" href="${escAttr(channel)}" target="_blank" rel="noopener">Go to our YouTube channel &rarr;</a>`
    : '';

  const only = (vids || []).filter(v => v.youtubeId);

  if (!only.length) {
    el.innerHTML =
      `<p class="vempty">Everything we publish is on the CheckFire channel.</p>` + channelLink;
    return;
  }

  // Newest first — "the most recent uploaded" is the whole point of
  // the box, so don't rely on an upstream sort.
  const sorted = [...only].sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
  const lead = sorted[0];
  // 1 Sep 2026 — with both thumbnails now the same size (marketing:
  // "can those thumbnails be the same size please") two videos is what
  // honestly fits. The hero box is height-capped by the grid, and a
  // third row was being clipped in half — which is what made the old
  // 52x34 chip look like a different kind of thing in the first place.
  // Everything else is one click away on the channel link below.
  const rest = sorted.slice(1, 2);

  // The thumbnail is an <img>, not a background image, so it can remove
  // ITSELF if it fails to load. The web filter is why the embedded
  // player had to go, and it may well block i.ytimg.com too — if it
  // does, the card falls back to the brand gradient underneath rather
  // than showing a broken image.
  const thumbImg = url => url
    ? `<img class="vthumb-img" src="${escAttr(safeUrl(url, ''))}" alt="" loading="lazy" onerror="this.remove()">`
    : '';

  // 1 Sep 2026 — marketing: "can those thumbnails be the same size
  // please".
  //
  // The box used to be a big 104px lead card with a 52x34 chip under
  // it, which read as two different kinds of thing stacked on top of
  // each other. It is one list now, every row identical, the newest
  // one flagged "Latest". That also fixes a quieter bug: the box is
  // height-capped by the hero grid, and the old pair of sizes
  // overflowed it, so the second video was always cut in half. Two
  // rows fit honestly; the rest are one click away on the channel.
  const row = (v, isLead) => `
      <a class="hbox-vid" href="${escAttr(safeUrl(v.href, channel || '#'))}" target="_blank" rel="noopener">
        <span class="hbox-vid-thumb">
          ${thumbImg(v.thumb)}
          <span class="vlead-play">
            <svg viewBox="0 0 24 24" fill="currentColor" width="11" height="11"><path d="M8 5v14l11-7z"/></svg>
          </span>
        </span>
        <span class="hbox-vid-title">${isLead ? '<b class="vrow-latest">Latest</b> ' : ''}${escHtml(v.title)}</span>
      </a>`;

  el.innerHTML = `
    <div class="vrest">${[lead].concat(rest).map((v, i) => row(v, i === 0)).join('')}</div>
    ${channelLink}`;
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

  renderNewsTicker(
    launches.status  === 'fulfilled' ? launches.value  : [],
    campaigns.status === 'fulfilled' ? campaigns.value : []
  );

  // Trade & Events is driven by the Documents/Events folders, not a list.
  await loadTradeEvents();
}

// "Latest updates" — a dock in the bottom-right corner.
//
// This was a full-width scrolling ticker across the foot of the home
// page. Marketing liked the idea but wanted it "more in the corner,
// with it popping up with the latest" (26 Aug), so it is now a small
// red pill that sits out of the way, pops itself open on load with the
// newest update, cycles gently through the rest, and expands to the
// full list when clicked.
//
// Still costs nothing extra — it reads the launches and campaigns
// loadSharePointData() has already fetched. Hidden when there is
// nothing to say. It lives on every page now, not just home: an
// update is worth seeing wherever you happen to be.
let _updates = [];
let _updIndex = 0;
let _updTimer = null;
let _updOpen  = false;

//
// 1 Sep 2026 — marketing: "How can we manage this? Ideally, I would
// like it to include only the most urgent updates."
//
// It was every launch and every campaign, newest eight — a feed, not
// something anyone decided. It now shows only the rows with the Yes/No
// column `Pinned` ticked (see HUB_CONFIG.updates). Tick two things and
// two things show; untick them and the dock is gone. Nothing ticked and
// the corner is empty, which is the true answer to "is anything
// urgent?" rather than eight things pretending to be.
function _updCol() {
  return (HUB_CONFIG.updates && HUB_CONFIG.updates.column) || 'Pinned';
}

function _updPinned(f) {
  const v = f[_updCol()];
  if (v === true) return true;
  if (typeof v === 'string') return /^(1|true|yes|y)$/i.test(v.trim());
  return v === 1;
}

// 1 Sep 2026 — David: "our breaking news box has disappeared."
//
// It had. Deck 5 changed the dock to show only rows with the Yes/No
// column ticked, and the column has not been added yet, so nothing was
// ever ticked and the dock hid itself exactly as designed. Designed
// wrong: "nobody has set this up" and "nothing is urgent today" look
// identical to the code and are completely different to the person
// looking at the page.
//
// So the two cases are told apart. If not one row even HAS the column,
// it doesn't exist — the hub carries on showing the newest few, as it
// did before. Once the column is there and someone has ticked nothing,
// that is a decision, and the corner stays empty.
function _updColumnExists(items) {
  const col = _updCol();
  return (items || []).some(f => typeof f[col] !== 'undefined' && f[col] !== null);
}

function renderNewsTicker(launches, campaigns) {
  const dock = document.getElementById('updates-dock');
  if (!dock) return;

  const cfg = HUB_CONFIG.updates || {};
  const items = [];

  (launches || []).forEach(f => {
    if (!f.Title) return;
    const when = fmtSpDate(f.LaunchDate);
    items.push({
      kind: 'launch',
      label: 'Launch',
      title: f.Title,
      text: `${f.Title}${f.Status ? ' — ' + f.Status : ''}`,
      when: when || 'Date to be confirmed',
      sort: String(f.LaunchDate || ''),
      pinned: _updPinned(f),
    });
  });

  (campaigns || []).forEach(f => {
    if (!f.Title) return;
    const span = [fmtSpDate(f.StartDate), fmtSpDate(f.EndDate)].filter(Boolean).join(' – ');
    items.push({
      kind: 'campaign',
      label: 'Campaign',
      title: f.Title,
      text: `${f.Title}${f.Status ? ' — ' + f.Status : ''}`,
      when: span,
      sort: String(f.StartDate || ''),
      pinned: _updPinned(f),
    });
  });

  const haveColumn = _updColumnExists([].concat(launches || [], campaigns || []));
  const pinned = items.filter(i => i.pinned);
  let shortlist = pinned;

  if (!pinned.length) {
    if (!haveColumn || cfg.requirePinned === false) {
      shortlist = items;
      console.info(`[Updates] no "${_updCol()}" column on the lists yet — showing the newest few. `
        + `Add a Yes/No column called "${_updCol()}" to Product Launches and Campaigns to choose what appears here.`);
    } else {
      console.info(`[Updates] the "${_updCol()}" column exists and nothing is ticked, so the dock stays hidden.`);
      dock.style.display = 'none';
      return;
    }
  }

  if (!shortlist.length) { dock.style.display = 'none'; return; }

  // Newest first.
  shortlist.sort((a, b) => b.sort.localeCompare(a.sort));
  _updates = shortlist.slice(0, cfg.max || 5);
  _updIndex = 0;
  dock.style.display = '';

  // Someone who has already dismissed it today shouldn't have it thrown
  // at them again on every page load — the pill still sits there, just
  // closed, with the count on it.
  const dismissed = (() => {
    try { return localStorage.getItem('cf-updates-seen') === _updStamp(); }
    catch (_) { return false; }
  })();

  _updRender();
  if (!dismissed) setTimeout(() => _updSetOpen(true), 1400);
}

// One "seen" stamp per day per newest-item, so a genuinely new update
// pops up again even if you dismissed yesterday's.
function _updStamp() {
  const newest = _updates[0] ? _updates[0].text : '';
  return new Date().toISOString().slice(0, 10) + '|' + newest.slice(0, 40);
}

function _updRender() {
  const dock = document.getElementById('updates-dock');
  if (!dock || !_updates.length) return;
  const n = _updates[_updIndex] || _updates[0];

  dock.className = 'upd-dock' + (_updOpen ? ' open' : '');
  dock.innerHTML = `
    <button class="upd-pill" onclick="toggleUpdates()" aria-expanded="${_updOpen}">
      <span class="upd-spark"></span>
      <span class="upd-pill-label">Latest updates</span>
      <span class="upd-count">${_updates.length}</span>
    </button>

    <div class="upd-panel" role="region" aria-label="Latest updates">
      <div class="upd-head">
        <span class="upd-head-title">Latest updates</span>
        <button class="upd-x" onclick="dismissUpdates()" aria-label="Close">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" width="14" height="14"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
      </div>

      <!-- Peek: one item at a time, cycling. This is what "pops up". -->
      <div class="upd-peek" id="upd-peek">
        ${_updItemHtml(n)}
      </div>

      <div class="upd-all" id="upd-all">
        ${_updates.map(_updItemHtml).join('')}
      </div>

      <button class="upd-more" id="upd-more" onclick="expandUpdates()">
        See all ${_updates.length} updates
      </button>
    </div>`;
}

// 1 Sep 2026 — marketing: "when we're announcing a new launch, to
// include a link to the folder containing all of the assets on the
// Product Launches or Campaigns page". That page already lists the real
// SharePoint folders for the item, so the update opens it rather than
// duplicating a link that could go stale.
function _updItemHtml(n) {
  if (!n) return '';
  const canOpen = (HUB_CONFIG.updates || {}).linkToAssets !== false && !!n.title;
  const inner = `
      <span class="upd-kind ${escAttr(n.kind)}">${escHtml(n.label)}</span>
      <div class="upd-text">${escHtml(n.text)}</div>
      ${n.when ? `<div class="upd-when">${escHtml(n.when)}</div>` : ''}
      ${canOpen ? `<div class="upd-go">Open the assets →</div>` : ''}`;
  if (!canOpen) return `<div class="upd-item">${inner}</div>`;
  const args = `'${escAttr(String(n.kind).replace(/'/g, ''))}','${escAttr(String(n.title).replace(/'/g, '&#39;'))}'`;
  return `<div class="upd-item is-link" role="button" tabindex="0"
       onclick="updOpenItem(${args})"
       onkeydown="if(event.key==='Enter')updOpenItem(${args})">${inner}</div>`;
}

// Opening an update takes you to the launch or campaign itself, where
// its asset folders are listed. The dock closes on the way so it isn't
// sat over the page you just asked for.
function updOpenItem(kind, title) {
  dismissUpdates();
  if (typeof srchOpenItem === 'function') srchOpenItem(kind, title);
}

function _updSetOpen(open) {
  _updOpen = !!open;
  const dock = document.getElementById('updates-dock');
  if (dock) dock.classList.toggle('open', _updOpen);
  clearInterval(_updTimer);
  if (_updOpen && _updates.length > 1) {
    // Cycle the peek line every 6s while the panel is open, unless the
    // user has expanded it to the full list.
    _updTimer = setInterval(() => {
      const dockEl = document.getElementById('updates-dock');
      if (!dockEl || dockEl.classList.contains('expanded')) return;
      _updIndex = (_updIndex + 1) % _updates.length;
      const peek = document.getElementById('upd-peek');
      if (!peek) return;
      peek.classList.add('swap');
      setTimeout(() => {
        peek.innerHTML = _updItemHtml(_updates[_updIndex]);
        peek.classList.remove('swap');
      }, 220);
    }, 6000);
  }
}

function toggleUpdates() {
  _updSetOpen(!_updOpen);
}

function expandUpdates() {
  const dock = document.getElementById('updates-dock');
  if (dock) dock.classList.add('expanded');
  clearInterval(_updTimer);
}

function dismissUpdates() {
  _updSetOpen(false);
  const dock = document.getElementById('updates-dock');
  if (dock) dock.classList.remove('expanded');
  try { localStorage.setItem('cf-updates-seen', _updStamp()); } catch (_) {}
}

// The corner dock is meant to be on EVERY page, but it only ever got
// rendered by loadSharePointData(), which runs on Launches/Campaigns.
// Land on the home page and there was nothing there until you'd been
// somewhere else first. This gives it its own loader; both lists are
// already in the 5-minute cache by the time anyone clicks through, so
// it costs nothing.
async function loadUpdatesDock() {
  if (window.HUB_DEMO_MODE) return;
  if (typeof getAccessToken !== 'function') return;
  try {
    const [launches, campaigns] = await Promise.all([
      fetchListItems(HUB_CONFIG.lists.launches).catch(() => []),
      fetchListItems(HUB_CONFIG.lists.campaigns).catch(() => []),
    ]);
    renderNewsTicker(launches, campaigns);
  } catch (e) {
    console.info('[Updates] unavailable:', e.message);
  }
}

// Which in-hub folder browsers have been started. Keyed by library, so
// 'product' and 'resources' each keep their own place in the tree.
const _fbLoaded = {};

// Resources used to BE the folder browser. It's now the library front
// door (loadResourcesLibrary), with the tree behind "Browse folders".
// Kept as an alias so ui.js's loadResourcesData() still resolves.
async function loadSharePointDocuments() {
  return loadResourcesLibrary();
}

// ═══ Library front doors ═════════════════════════════════════
//
// GENERALISED 26 Aug 2026. This started as the Product Portal front
// door and now drives Resources as well, so the two pages look and
// behave the same — David: "Resources and documents needs to fit more
// in to the rest of the website".
//
// Both sites are folder trees of files. The hub reads the whole tree
// once, tags every file, and offers three ways in — search, a tile
// row, and type chips — with Download and Copy link on every row so
// "staff have everything all in one place".
//
// Nothing in SharePoint changes. Marketing keep filing as they do.

const LIB = {};   // key → { files, loaded, driveId, tag, cat, q }

function _libCfg(key) {
  return ((HUB_CONFIG.libraries || {})[key]) || {};
}

// Resolve "01. Marketing/08. PDF PIF, Data Sheets…" to a folder id,
// one forgiving step at a time. Returns null (and says why in the
// console) rather than throwing, so one renamed folder never costs you
// the other six sources.
async function _libResolvePath(driveId, path) {
  let id = null;
  for (const part of String(path).split('/').map(p => p.trim()).filter(Boolean)) {
    const hit = await _findChildFolder(driveId, id, part);
    if (!hit) { console.info(`[Library] no folder "${part}" under ${path} — skipping that root.`); return null; }
    id = hit.id;
  }
  return id;
}

async function _libCrawl(driveId, itemId, path, depth, out, cap, exclude) {
  if (depth < 0 || out.length >= cap) return;
  let kids;
  try { kids = await fetchDriveChildren(driveId, itemId); }
  catch (_) { return; }

  const folders = [];
  for (const k of kids) {
    if (k.folder) {
      // Folders that have their own page in the hub aren't repeated here.
      if (!path.length && exclude.some(x => x.toLowerCase() === String(k.name).toLowerCase())) continue;
      folders.push(k);
    } else if (out.length < cap) {
      out.push(Object.assign({}, k, { _path: path }));
    }
  }
  await Promise.all(folders.map(f =>
    _libCrawl(driveId, f.id, path.concat(f.name), depth - 1, out, cap, exclude)));
}

// A category used to be "whichever folder it sits in". With several
// sites feeding one page that stops working — the same kind of document
// lives under a different folder name on every site. A rule can now
// carry `match`, a pattern tried against the file's whole path and
// name. Folder rules are still tried first, so nothing regresses.
function _libCatLabel(key, folder, f) {
  const rows = _libCfg(key).categories || [];

  const hit = rows.find(r => r.folder &&
    String(r.folder).toLowerCase() === String(folder || '').toLowerCase());
  if (hit) return hit.label;

  if (f) {
    // The FILE NAME is asked first, and the folder path only if the name
    // says nothing. Both at once gets it wrong in a way that matters:
    // sales.marketing files PIFs, datasheets and MSDS together in
    // "08. PDF PIF, Data Sheets, MSDS Sheets & Toolkits", so a path
    // match would file "Commander PIF.pdf" under MSDS purely because
    // the word appears in the folder name three levels up.
    const name = String(f.name || '').toLowerCase();
    const path = [].concat(f._path || []).join(' / ').toLowerCase();
    for (const hay of [name, path]) {
      if (!hay) continue;
      for (const r of rows) {
        if (!r.match) continue;
        try { if (new RegExp(r.match, 'i').test(hay)) return r.label; }
        catch (_) { /* a bad pattern in config mustn't break the page */ }
      }
    }
  }
  return folder || 'General';
}

function _libTag(key, f) {
  const rows = _libCfg(key).tags || [];
  const hay = [f.name].concat(f._path || []).join(' ').toLowerCase();
  for (const r of rows) {
    try { if (new RegExp(r.match, 'i').test(hay)) return r; }
    catch (_) { /* a bad pattern in config mustn't break the page */ }
  }
  return null;
}

async function loadLibrary(key) {
  const cfg  = _libCfg(key);
  const host = document.getElementById(cfg.hostId);
  if (!host) return;

  const signedIn = window.AUTH && window.AUTH.account;
  if (window.HUB_DEMO_MODE || !signedIn) {
    host.innerHTML = '<p class="prose dim">Sign in with your CheckFire account to open this library.</p>';
    return;
  }

  if (LIB[key] && LIB[key].loaded) { renderLibrary(key); return; }
  LIB[key] = { files: [], loaded: false, driveId: null, tag: 'all', cat: 'all', q: '' };

  host.innerHTML = `<div class="lib-boot">
    <div class="skeleton sk-line med"></div>
    <div class="skeleton sk-line"></div>
    <div class="skeleton sk-line short"></div>
    <p class="prose dim" style="margin-top:12px">Reading ${escHtml(cfg.title || 'the library')}…</p>
  </div>`;

  try {
    const site  = cfg.site === 'product' ? HUB_CONFIG.productPortalSite : HUB_CONFIG.sharepointSite;

    // One source, or several. `sources` is what turns the Product
    // Portal page into the home for everything (1 Sep 2026, round 2);
    // without it this behaves exactly as it did.
    const sources = (cfg.sources && cfg.sources.length)
      ? cfg.sources
      : [{ key: 'main', label: '', site, library: cfg.library || HUB_CONFIG.documentsLibrary,
           depth: cfg.crawlDepth || 3, max: cfg.maxFiles || 400 }];

    const out = [];
    const tally = [];

    await Promise.all(sources.map(async src => {
      const before = out.length;
      try {
        const drives = src.allLibraries
          ? await resolveAllDrives(src.site)
          : [await resolveDrive(src.site, src.library || HUB_CONFIG.documentsLibrary)];

        if (!LIB[key].driveId && drives[0]) LIB[key].driveId = drives[0].id;

        for (const drive of drives) {
          const mine = [];
          const cap  = src.max || cfg.maxFiles || 400;

          if (src.roots && src.roots.length) {
            for (const root of src.roots) {
              if (mine.length >= cap) break;
              const id = await _libResolvePath(drive.id, root);
              if (!id) continue;
              // The root's own name leads the path, so it still reads
              // as "where this came from" in the file list.
              await _libCrawl(drive.id, id, [String(root).split('/').pop()],
                              (src.depth || 3), mine, cap, src.excludeFolders || []);
            }
          } else {
            await _libCrawl(drive.id, null, [], (src.depth || 3), mine, cap,
                            src.excludeFolders || cfg.excludeFolders || []);
          }

          mine.forEach(f => {
            f._source    = src.label || '';
            f._sourceKey = src.key || '';
            // A system library's name is worth keeping — it is how the
            // Product Change Notifications turn up at all.
            if (drives.length > 1 && drive.name && !/^documents$/i.test(drive.name)) {
              f._library = drive.name;
              if (!(f._path || []).length) f._path = [drive.name];
            }
          });
          out.push(...mine);
        }
      } catch (e) {
        console.info(`[Library] source "${src.label || src.key}" unavailable: ${e.message}`);
      }
      tally.push(`${src.label || src.key}: ${out.length - before}`);
    }));

    console.info('[Library] ' + (cfg.title || key) + ' — ' + tally.join(' · '));

    // The same document is filed on more than one site. Keep the first
    // and remember there was another, rather than listing it twice.
    const seen = new Set();
    const merged = [];
    out.forEach(f => {
      const sig = _slugKey(f.name) + '|' + (f.size || 0);
      if (seen.has(sig)) return;
      seen.add(sig);
      merged.push(f);
    });
    if (merged.length !== out.length) {
      console.info(`[Library] ${out.length - merged.length} duplicate file(s) across sites folded together.`);
    }

    LIB[key].files = merged.map(f => {
      const t = _libTag(key, f);
      return Object.assign({}, f, {
        _tag:      t ? t.key : 'other',
        _tagLbl:   t ? t.label : 'Other',
        _catFolder: (f._path || [])[0] || '',
        _cat:      _libCatLabel(key, (f._path || [])[0], f),
        _sub:      [(f._path || [])[1] || '', f._source || ''].filter(Boolean).join(' · '),
      });
    });
    LIB[key].loaded = true;
    renderLibrary(key);
  } catch (e) {
    const msg = e.message === 'NOT_FOUND'
      ? 'That SharePoint site or library could not be found — check the URL in config.js and that you have access to it.'
      : `Couldn't read the library: ${e.message}`;
    host.innerHTML = `<p class="sp-error">${escHtml(msg)}</p>`;
  }
}

function renderLibrary(key) {
  const cfg   = _libCfg(key);
  const host  = document.getElementById(cfg.hostId);
  const state = LIB[key];
  if (!host || !state) return;

  const files = state.files;
  if (!files.length) {
    host.innerHTML = '<p class="prose dim">Nothing in this library yet.</p>';
    return;
  }

  // 1 Sep 2026 — marketing on the Resources page: "This feels very
  // chaotic at the moment. Can we just have a list as before with
  // documents added to the SharePoint."
  //
  // Fair. Three files were being presented through a tile row, a chip
  // row and a "recently updated" rail — three navigation devices over
  // a list you could read in one glance. In simple mode there is a
  // search box and the documents, grouped by the folder they live in,
  // and that is all. Set `simple:false` in config.js to get the
  // faceted view back; the Product Portal, which has 33 files across
  // ten products, still uses it.
  if (_libCfg(key).simple) {
    host.innerHTML = `
      <div class="lib-search-wrap">
        <svg class="lib-search-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" width="18" height="18"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
        <input class="lib-search" id="lib-q-${escAttr(key)}" type="search" autocomplete="off"
               placeholder="${escAttr(cfg.searchPlaceholder || 'Search everything here…')}"
               oninput="libSearch('${escAttr(key)}',this.value)">
        <span class="lib-search-count">${files.length} file${files.length === 1 ? '' : 's'}</span>
      </div>
      <div class="lib-results" id="lib-results-${escAttr(key)}"></div>`;
    renderLibraryResults(key);
    return;
  }

  const tiles = (cfg.tags || []).map(t => ({
    key: t.key, label: t.label,
    n: files.filter(f => f._tag === t.key).length,
  })).filter(t => t.n);
  const otherN = files.filter(f => f._tag === 'other').length;
  if (tiles.length && otherN) tiles.push({ key: 'other', label: 'Other', n: otherN });

  const catOrder = (cfg.categories || []).map(c => c.label);
  const cats = [...new Set(files.map(f => f._cat))].sort((a, b) => {
    const ia = catOrder.indexOf(a), ib = catOrder.indexOf(b);
    return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || a.localeCompare(b);
  }).map(label => ({ label, n: files.filter(f => f._cat === label).length }));

  const recent = [...files]
    .sort((a, b) => String(b.lastModifiedDateTime || '').localeCompare(String(a.lastModifiedDateTime || '')))
    .slice(0, cfg.recentCount || 6);

  host.innerHTML = `
    <div class="lib-search-wrap">
      <svg class="lib-search-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" width="18" height="18"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
      <input class="lib-search" id="lib-q-${escAttr(key)}" type="search" autocomplete="off"
             placeholder="${escAttr(cfg.searchPlaceholder || 'Search everything here…')}"
             oninput="libSearch('${escAttr(key)}',this.value)">
      <span class="lib-search-count">${files.length} files</span>
    </div>

    ${tiles.length ? `
    <div class="lib-sec-head"><h2 class="lib-sec-title">${escHtml(cfg.tagsLabel || 'By product')}</h2>
      <button class="lib-reset" onclick="libReset('${escAttr(key)}')">Reset</button></div>
    <div class="lib-tiles" id="lib-tiles-${escAttr(key)}">
      ${tiles.map((t, i) => `
        <button class="lib-tile" style="--i:${i}" onclick="libPick('${escAttr(key)}','${escAttr(t.key)}',this)">
          <span class="lib-tile-n">${t.n}</span>
          <span class="lib-tile-l">${escHtml(t.label)}</span>
        </button>`).join('')}
    </div>` : ''}

    <div class="lib-sec-head"><h2 class="lib-sec-title">${escHtml(cfg.catsLabel || 'By type')}</h2>
      ${tiles.length ? '' : `<button class="lib-reset" onclick="libReset('${escAttr(key)}')">Reset</button>`}</div>
    <div class="lib-cats" id="lib-cats-${escAttr(key)}">
      <button class="lib-cat active" onclick="libCat('${escAttr(key)}','all',this)">All<b>${files.length}</b></button>
      ${cats.map(c => `<button class="lib-cat" onclick="libCat('${escAttr(key)}','${escAttr(c.label)}',this)">${escHtml(c.label)}<b>${c.n}</b></button>`).join('')}
    </div>

    ${recent.length ? `
    <div class="lib-recent">
      <div class="lib-recent-lbl">Recently updated</div>
      <div class="lib-recent-row">
        ${recent.map(f => {
          const k = regDoc(f);
          return `<button class="lib-recent-card" onclick="openRegDoc('${k}')">
            <span class="lib-recent-name">${escHtml(String(f.name).replace(/\.[a-z0-9]+$/i, ''))}</span>
            <span class="lib-recent-meta">${escHtml(f._cat)} · ${escHtml(fmtSpDate(f.lastModifiedDateTime))}</span>
          </button>`;
        }).join('')}
      </div>
    </div>` : ''}

    <div class="lib-results" id="lib-results-${escAttr(key)}"></div>`;

  renderLibraryResults(key);
}

function renderLibraryResults(key) {
  const cfg   = _libCfg(key);
  const state = LIB[key];
  const box   = document.getElementById('lib-results-' + key);
  if (!box || !state) return;

  const q = state.q.trim().toLowerCase();
  // state.cat is 'all', one category label, or a list of them — a
  // Product Portal band can cover several ("Data sheets" + "MSDS").
  const wantCat = f => state.cat === 'all'
    || (Array.isArray(state.cat) ? state.cat.indexOf(f._cat) >= 0 : f._cat === state.cat);

  const rows = state.files.filter(f =>
    (state.tag === 'all' || f._tag === state.tag) &&
    wantCat(f) &&
    (!q || (f.name + ' ' + (f._path || []).join(' ') + ' ' + (f._source || '')).toLowerCase().includes(q)));

  const filtered = state.tag !== 'all' || state.cat !== 'all' || !!q;

  if (!rows.length) {
    box.innerHTML = `<div class="px-empty"><h3>Nothing matches</h3>
      <p>No files for that combination.
      <button class="lib-reset inline" onclick="libReset('${escAttr(key)}')">Clear the filters</button> and try again.</p></div>`;
    return;
  }

  const groups = {};
  rows.forEach(f => { (groups[f._cat] = groups[f._cat] || []).push(f); });

  const catOrder = (cfg.categories || []).map(c => c.label);
  const keys = Object.keys(groups).sort((a, b) => {
    const ia = catOrder.indexOf(a), ib = catOrder.indexOf(b);
    return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || a.localeCompare(b);
  });

  box.innerHTML = `
    <div class="lib-results-head">
      <span>${rows.length} file${rows.length === 1 ? '' : 's'}${filtered ? ' matching' : ''}</span>
      ${filtered ? `<button class="lib-reset inline" onclick="libReset('${escAttr(key)}')">Clear filters</button>` : ''}
    </div>
    ${keys.map(k => `
      <section class="lib-group">
        <h3 class="lib-group-head">${escHtml(k)}<span>${groups[k].length}</span></h3>
        <div class="lib-files">
          ${groups[k]
            .sort((a, b) => String(a.name).localeCompare(String(b.name)))
            .map(f => libFileRow(f)).join('')}
        </div>
      </section>`).join('')}`;
}

// One file row.
//
// 1 Sep 2026 (round 2) — clicking a row used to navigate you away to
// the reader page. David: "I want it to open everything you click on it
// and it brings everything open." So a click opens the document HERE,
// underneath its own row, and clicking again folds it away. The reader
// is still one button along for anyone who wants the whole screen.
function libFileRow(f, subLabel) {
  const k = regDoc(f);
  const meta = [subLabel || f._sub, f._tagLbl && f._tagLbl !== 'Other' ? f._tagLbl : '',
                humanSize(f.size), fmtSpDate(f.lastModifiedDateTime)]
    .filter(Boolean).map(escHtml).join(' · ');
  return `
    <div class="lib-row" id="lr-${k}">
      <div class="lib-file" role="button" tabindex="0" aria-expanded="false"
           onclick="libToggleFile('${k}')" onkeydown="if(event.key==='Enter')libToggleFile('${k}')">
        <span class="lib-file-ico">${escHtml((String(f.name).split('.').pop() || 'FILE').slice(0, 4).toUpperCase())}</span>
        <span class="lib-file-main">
          <span class="lib-file-name">${escHtml(String(f.name).replace(/\.[a-z0-9]+$/i, ''))}</span>
          <span class="lib-file-meta">${meta}</span>
        </span>
        ${docActions(k)}
      </div>
      <div class="lib-open" id="lo-${k}"></div>
    </div>`;
}

// Open the document in the row, or fold it away again. Reuses the same
// renderer the campaign pages use, so a PDF looks the same wherever you
// meet it in the hub.
function libToggleFile(k) {
  const host = document.getElementById('lo-' + k);
  const row  = document.getElementById('lr-' + k);
  const head = row && row.querySelector('.lib-file');
  if (!host) { openRegDoc(k); return; }

  if (host.firstChild) {
    host.innerHTML = '';
    if (row)  row.classList.remove('open');
    if (head) head.setAttribute('aria-expanded', 'false');
    return;
  }

  const f = DOCREG[k];
  if (!f) return;
  if (row)  row.classList.add('open');
  if (head) head.setAttribute('aria-expanded', 'true');

  host.innerHTML = `<div class="oe-stage" data-doc="${k}" data-kind="${_oeKind(f.name)}">
      <div class="oe-wait"><span class="rdr-spin"></span></div>
    </div>`;
  _oeFill(host.firstElementChild);
}

function libPick(key, tag, btn) {
  const s = LIB[key];
  if (!s) return;
  s.tag = (s.tag === tag) ? 'all' : tag;
  document.querySelectorAll('#lib-tiles-' + key + ' .lib-tile')
    .forEach(b => b.classList.toggle('active', b === btn && s.tag === tag));
  renderLibraryResults(key);
}

function libCat(key, label, btn) {
  const s = LIB[key];
  if (!s) return;
  s.cat = label;
  document.querySelectorAll('#lib-cats-' + key + ' .lib-cat')
    .forEach(b => b.classList.toggle('active', b === btn));
  renderLibraryResults(key);
}

const _libTimers = {};
function libSearch(key, v) {
  const s = LIB[key];
  if (!s) return;
  s.q = v || '';
  clearTimeout(_libTimers[key]);
  _libTimers[key] = setTimeout(() => renderLibraryResults(key), 140);
}

function libReset(key) {
  const s = LIB[key];
  if (!s) return;
  s.tag = 'all'; s.cat = 'all'; s.q = '';
  const q = document.getElementById('lib-q-' + key);
  if (q) q.value = '';
  document.querySelectorAll('#lib-tiles-' + key + ' .lib-tile').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('#lib-cats-' + key + ' .lib-cat').forEach((b, i) => b.classList.toggle('active', i === 0));
  renderLibraryResults(key);
}

// ── The two pages that use it ─────────────────────────────────

// ── Product Portal front door ─────────────────────────────────
//
// 1 Sep 2026. The product team asked for the whole Product Portal to
// live in the hub, "ideally keeping the organisation the same", and
// listed what the hub's page was missing: product change
// notifications, links to PIF, the sample request sheet, the new
// product request sheet, a launch countdown and upcoming dates, a
// feedback form and datasheets/MSDS.
//
// The certificate index below is untouched — same products, same
// document types, same search box. These are three bands around it,
// and every one of them is driven by something that either exists or
// doesn't: a folder, a list, a URL in config.js. Nothing renders a
// placeholder. If Aneta and Jess haven't made the folder yet, the band
// isn't there, and the page still reads as finished.
async function loadProductPortal() {
  const done = loadLibrary('product');
  renderPortalUpcoming();      // the launches list, not the library
  await done;
  renderPortalSections();
  renderPortalFeedback();
}

// "Launch countdown and upcoming dates to look out for."
async function renderPortalUpcoming() {
  const host = document.getElementById('pp-upcoming');
  if (!host) return;
  host.innerHTML = '';
  if (typeof getAccessToken !== 'function' || window.HUB_DEMO_MODE) return;

  let items = [];
  try { items = await fetchListItems(HUB_CONFIG.lists.launches); }
  catch (e) { console.info('[Portal] launches not read:', e.message); return; }

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const rows = (items || [])
    .map(f => ({ title: f.Title, status: f.Status, raw: f.LaunchDate,
                 // _toIsoDate lives in jotform.js, which loads after this
                 // file — guard the reference rather than assume the order.
                 when: (typeof _toIsoDate === 'function' ? _toIsoDate(f.LaunchDate) : '') || '' }))
    .filter(r => r.title && r.when && new Date(r.when + 'T00:00:00') >= today)
    .sort((a, b) => a.when.localeCompare(b.when));

  if (!rows.length) return;

  const cfg  = HUB_CONFIG.productPortal || {};
  const next = rows[0];
  const days = Math.round((new Date(next.when + 'T00:00:00') - today) / 86400000);
  const rest = rows.slice(1, (cfg.upcomingCount || 4));

  host.innerHTML = `
    <div class="pp-band">
      <div class="pp-band-head">
        <h2 class="pp-band-title">Coming up</h2>
        <span class="pp-band-note">From the Product Launches list</span>
      </div>
      <div class="pp-next">
        <div class="pp-count">
          <b>${days === 0 ? 'Today' : days}</b>
          <span>${days === 0 ? 'launching' : days === 1 ? 'day to go' : 'days to go'}</span>
        </div>
        <div>
          <p class="pp-next-title">${escHtml(next.title)}</p>
          <div class="pp-next-meta">${escHtml(fmtSpDate(next.raw))}${next.status ? ' · ' + escHtml(next.status) : ''}</div>
        </div>
      </div>
      ${rest.length ? `<div class="pp-dates">${rest.map(r => `
        <span class="pp-date"><span class="pp-dot"></span><b>${escHtml(r.title)}</b> ${escHtml(fmtSpDate(r.raw))}</span>`).join('')}</div>` : ''}
    </div>`;
}

// The extra document groups. A band exists only if its folder does —
// and the folders are already in the crawl the library just did, so
// this costs no extra Graph calls.
function _ppMatchFolder(sec, folderName) {
  const key = _slugKey(folderName);
  return [sec.folder].concat(sec.aliases || []).some(x => _slugKey(x) === key);
}

function renderPortalSections() {
  const host = document.getElementById('pp-sections');
  if (!host) return;
  host.innerHTML = '';

  const state = LIB.product;
  const secs  = (HUB_CONFIG.productPortal && HUB_CONFIG.productPortal.sections) || [];
  if (!state || !state.loaded || !secs.length) return;

  const folders = [...new Set(state.files.map(f => f._catFolder).filter(Boolean))];

  const live = [];
  secs.forEach(sec => {
    // By kind first — that is what survives files coming from five
    // different sites with five different folder conventions.
    if (sec.cats && sec.cats.length) {
      const present = sec.cats.filter(c => state.files.some(f => f._cat === c));
      const count   = state.files.filter(f => present.indexOf(f._cat) >= 0).length;
      if (count) { live.push({ sec, count, cat: present }); return; }
    }
    // Then by folder, for the things that are a folder rather than a
    // kind of document — the two request sheets.
    const folder = folders.find(fn => _ppMatchFolder(sec, fn));
    if (!folder) return;
    const count = state.files.filter(f => f._catFolder === folder).length;
    if (!count) return;
    live.push({ sec, count, cat: _libCatLabel('product', folder) });
  });

  if (!live.length) {
    console.info('[Portal] nothing matched a band yet — bands hidden.');
    return;
  }

  PP_BANDS = live;
  host.innerHTML = `
    <div class="pp-band">
      <div class="pp-band-head">
        <h2 class="pp-band-title">Everything, by what it is</h2>
        <span class="pp-band-note">${state.files.length} documents from ${_ppSourceCount(state)} SharePoint sources</span>
      </div>
      <div class="pp-secs">
        ${live.map((l, i) => `
          <button class="pp-sec" onclick="ppOpenSection(${i})">
            <span class="pp-sec-n">${l.count}</span>
            <span class="pp-sec-l">${escHtml(l.sec.label)}</span>
            <span class="pp-sec-d">${escHtml(l.sec.desc || '')}</span>
          </button>`).join('')}
      </div>
    </div>`;
}

let PP_BANDS = [];

function _ppSourceCount(state) {
  return new Set(state.files.map(f => f._source || 'Product Portal')).size;
}

// Opening a band is a filter over the index that is already on the
// page, not a second place for the same files to live.
function ppOpenSection(i) {
  const s = LIB.product;
  const band = PP_BANDS[i];
  if (!s || !band) return;
  s.tag = 'all'; s.q = ''; s.cat = band.cat;

  const q = document.getElementById('lib-q-product');
  if (q) q.value = '';
  document.querySelectorAll('#lib-tiles-product .lib-tile').forEach(b => b.classList.remove('active'));
  const wanted = Array.isArray(band.cat) ? band.cat : [band.cat];
  document.querySelectorAll('#lib-cats-product .lib-cat').forEach(b => {
    b.classList.toggle('active', wanted.some(c => b.textContent.indexOf(c) === 0));
  });
  document.querySelectorAll('#pp-sections .pp-sec').forEach((b, bi) => b.classList.toggle('active', bi === i));

  renderLibraryResults('product');
  const box = document.getElementById('lib-results-product');
  if (box) box.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// "Also has a feedback form." One URL in config.js; no URL, no box.
function renderPortalFeedback() {
  const host = document.getElementById('pp-feedback');
  if (!host) return;
  const cfg = HUB_CONFIG.productPortal || {};
  if (!cfg.feedbackUrl) { host.innerHTML = ''; return; }
  host.innerHTML = `
    <div class="pp-fb">
      <div class="pp-fb-copy">
        <p class="pp-fb-title">${escHtml(cfg.feedbackTitle || 'Feedback on a product')}</p>
        <p class="pp-fb-sub">${escHtml(cfg.feedbackSub || '')}</p>
      </div>
      <a class="px-cta" href="${escAttr(safeUrl(cfg.feedbackUrl, '#'))}" target="_blank" rel="noopener">
        Send feedback
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
      </a>
    </div>`;
}
async function loadResourcesLibrary() { return loadLibrary('resources'); }

// "Browse folders" — the original folder tree, unchanged, for anyone
// who wants SharePoint's own structure. Works on both pages.
async function toggleLibraryBrowse(key, btn) {
  const cfg = _libCfg(key);
  const idx = document.getElementById(cfg.hostId);
  const br  = document.getElementById(cfg.browserId);
  if (!idx || !br) return;

  const showBrowser = br.style.display === 'none' || !br.style.display;
  br.style.display  = showBrowser ? '' : 'none';
  idx.style.display = showBrowser ? 'none' : '';
  if (btn) btn.textContent = showBrowser ? (cfg.backLabel || 'Back') : 'Browse folders';

  if (showBrowser && !_fbLoaded[key]) {
    _fbLoaded[key] = true;
    const site = cfg.site === 'product' ? HUB_CONFIG.productPortalSite : HUB_CONFIG.sharepointSite;
    await fbInit(key, site, cfg.library || HUB_CONFIG.documentsLibrary,
                 cfg.browserGridId, cfg.crumbId, cfg.title || 'Library');
  } else if (showBrowser) {
    renderBrowser(key);
  }
}

function togglePortalBrowse(btn)    { return toggleLibraryBrowse('product', btn); }
function toggleResourcesBrowse(btn) { return toggleLibraryBrowse('resources', btn); }

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

  // Folders drill in; files open in the reader and carry Download and
  // Copy link, so nothing has to be opened first to be shared.
  grid.innerHTML = `<div class="lib-files">${sorted.map(f => {
    const idx = b.items.indexOf(f);

    if (f.folder) {
      const n = (f.folder && f.folder.childCount) || 0;
      return `
        <div class="lib-file folder" role="button" tabindex="0"
             onclick="fbOpenFolder('${key}',${idx})" onkeydown="if(event.key==='Enter')fbOpenFolder('${key}',${idx})">
          <span class="lib-file-ico folder">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="17" height="17"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
          </span>
          <span class="lib-file-main">
            <span class="lib-file-name">${escHtml(f.name)}</span>
            <span class="lib-file-meta">${n} item${n === 1 ? '' : 's'}</span>
          </span>
          <span class="lib-file-go">Open →</span>
        </div>`;
    }

    if (!(f._driveId && f.id)) {
      return `<a class="lib-file" ${f.webUrl ? `href="${escAttr(safeUrl(f.webUrl))}" target="_blank" rel="noopener"` : ''}>
        <span class="lib-file-ico">${escHtml((String(f.name).split('.').pop() || 'FILE').slice(0, 4).toUpperCase())}</span>
        <span class="lib-file-main"><span class="lib-file-name">${escHtml(f.name)}</span></span>
      </a>`;
    }
    return libFileRow(f);
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


// ── Launch / campaign detail ──────────────────────────────────
//
// REBUILT 26 Aug 2026 (round 2 fix). David: "Nothing is opening as
// should on Product launches or Campaigns. not like I asked for."
// He was right — the index pages had been rebuilt as a website but
// clicking a card still opened the old dark panel with a grey grid of
// blocks under it. The detail now uses the SAME editorial shape as the
// page it came from: a lead spread with the item's own artwork, then
// its assets as real, openable, downloadable files.
//
// The metrics band ("Emails sent 0 / Social media posts 0 / Blogs 0 /
// PR activity 0") is GONE. Nothing feeds those columns and there is no
// plan to, so it was four zeroes taking up the width of the page.
// David: "Not sure how we can track this so think we should remove
// it!!" If tracking ever arrives, put it back deliberately.

function _renderDetail(opts) {
  // opts: { containerId, hideIds, item, kind, folderRoot, backLabel, backFn }
  const box = document.getElementById(opts.containerId);
  if (!box) return;
  const f = opts.item || {};

  (opts.hideIds || []).forEach(id => { const el = document.getElementById(id); if (el) el.style.display = 'none'; });
  box.style.display = '';
  window.scrollTo(0, 0);

  const isLaunch = opts.kind === 'launch';
  const landing  = safeUrl(linkOf(f.LinkURL), '');
  const dates    = [fmtSpDate(f.StartDate || f.LaunchDate), fmtSpDate(f.EndDate)].filter(Boolean).join(' – ');
  const codes    = isLaunch ? productCodes(f) : [];
  // Only a REAL description goes under the title. CampaignType is
  // already the eyebrow — repeating it as body copy ("BRAND / … /
  // Brand") read as leftover test text.
  const sub      = f.Description || f.Summary || '';
  const channels = Array.isArray(f.Channels) ? f.Channels : String(f.Channels || '').split(/[,;/]+/);
  const chips    = channels.map(c => String(c).trim()).filter(Boolean);

  // NO config placeholder blocks. 26 Aug 2026 (fix 4): the page used to
  // render HUB_CONFIG.campaignAssetBlocks straight away — six tiles
  // ("Infographic", "Email signature", "Data card"…) named after folders
  // that don't exist in SharePoint. They looked real, said "Open", and
  // did nothing when clicked, which is exactly what David reported:
  // "Campaigns doesn't open anything". The tiles are now built from the
  // folders that are actually there, and nothing is drawn until they are.

  box.innerHTML = `
    <div class="dt-backbar">
      <button class="dt-back" onclick="${opts.backFn}">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><polyline points="15 18 9 12 15 6"/></svg>
        ${escHtml(opts.backLabel)}
      </button>
    </div>

    <section class="px-lead dark">
      <div class="px-lead-copy">
        <div class="px-eyebrow">${escHtml(isLaunch ? 'Product launch' : (f.CampaignType || 'Campaign'))}</div>
        <h1 class="px-lead-title">${escHtml(f.Title || 'Untitled')}</h1>
        ${sub ? `<p class="px-lead-sub">${escHtml(sub)}</p>` : ''}
        ${codes.length ? `<div class="px-lead-codes">${codes.map(c => `<span class="px-code">${escHtml(c)}</span>`).join('')}</div>` : ''}
        ${!isLaunch && chips.length ? `<div class="px-lead-codes">${chips.map(c => `<span class="px-chan">${escHtml(c)}</span>`).join('')}</div>` : ''}
        <div class="px-lead-meta">
          <span class="px-badge"><span class="px-badge-dot ${ragOf(f.Status)}"></span>${escHtml(f.Status || 'Not set')}</span>
          ${dates ? `<span class="px-lead-when">${escHtml(dates)}</span>` : ''}
          ${f.Region ? `<span class="px-lead-when">${escHtml(f.Region)}</span>` : ''}
        </div>
        <div class="dt-actions">
          ${landing ? `<a class="px-cta" href="${escAttr(landing)}" target="_blank" rel="noopener">
            View the landing page
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
          </a>` : ''}
          ${isLaunch ? `<button class="px-cta ghost" onclick="openProductPortal()">Certificates &amp; datasheets</button>` : ''}
        </div>
      </div>
      <div class="px-lead-media" id="dt-hero-img">
        <span class="px-lead-initials">${escHtml(_pxInitials(f.Title))}</span>
      </div>
    </section>

    <div class="dt-main">
      <div class="dt-sec-head">
        <h2 class="dt-sec-title">Everything for this ${isLaunch ? 'launch' : 'campaign'}</h2>
        <p class="dt-sec-sub">All of it, open on the page — artwork, data sheets, decks and video. Download or copy a link on anything you need to send on.</p>
      </div>
      <!-- 1 Sep 2026 (round 2). This used to be a grid of folder tiles
           that opened a file list that opened a reader. Three clicks to
           see one document. It is the documents themselves now. -->
      <div id="cd-blocks"></div>
      <div id="cd-asset-panel"></div>
    </div>`;

  // Remember what the assets should resolve against.
  _detailContext = { folderRoot: opts.folderRoot, campaignFolder: f.CampaignFolder || f.Folder || f.Title };
  _detailBlocks  = [];
  _detailFolder  = null;

  _loadDetailAssets();
}

// Folder tiles. Each one is a real folder in SharePoint with a live
// file count; clicking it lists the files underneath.
function _blocksHtml(blocks) {
  return (blocks || []).map((bl, bi) => `
      <button class="dt-folder" onclick="openDetailAsset(${bi})">
        <span class="dt-folder-ico">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" width="19" height="19"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
        </span>
        <span class="dt-folder-main">
          <span class="dt-folder-name">${escHtml(bl.label)}</span>
          <span class="dt-folder-note">${bl.count === undefined ? 'Open' : `${bl.count} file${bl.count === 1 ? '' : 's'}`}</span>
        </span>
      </button>`).join('');
}

// The item's own artwork, from its asset folder — same source the
// cards use, so the detail page matches the card you clicked.
// Resolve  Documents/<root>/<item>  with a console trail, because when
// this misses the page has nothing to show and the old code failed
// SILENTLY — it just left the config placeholders sitting there looking
// like real folders. Every step now says what it found.
//
// The exact-name walk is the fast path. When it misses (a list Title
// that has drifted from the folder name, an extra word, a stray
// apostrophe) it falls back to a drive-wide search for the name and
// keeps only folders whose parent really is <root>.
async function _resolveItemFolder(driveId, rootName, itemName) {
  const root = await _findChildFolder(driveId, null, rootName);
  if (!root) {
    console.warn(`[Assets] there is no “${rootName}” folder at the top of the document library.`);
    return null;
  }

  const item = await _findChildFolder(driveId, root.id, itemName);
  if (item) return item;

  console.warn(`[Assets] no folder called “${itemName}” inside “${rootName}” — searching instead.`);
  try {
    const q = encodeURIComponent(String(itemName).slice(0, 60).replace(/'/g, "''"));
    const res = await graphFetch(
      `/drives/${driveId}/root/search(q='${q}')?$select=id,name,folder,parentReference&$top=50`);
    const tail = '/' + String(rootName).toLowerCase();
    const hits = (res.value || []).filter(x => x.folder &&
      String((x.parentReference && x.parentReference.path) || '').toLowerCase().endsWith(tail));
    if (hits.length) {
      const key = _slugKey(itemName) || '';
      hits.sort((a, b) =>
        Math.abs((_slugKey(a.name) || '').length - key.length) -
        Math.abs((_slugKey(b.name) || '').length - key.length));
      console.info(`[Assets] search matched “${hits[0].name}”.`);
      return Object.assign({}, hits[0], { _driveId: driveId });
    }
    console.warn(`[Assets] search found nothing under “${rootName}” for “${itemName}”.`);
  } catch (err) {
    console.warn('[Assets] folder search failed:', err.message);
  }
  return null;
}

// One pass: find the item's folder, then use it for BOTH the hero
// artwork and the asset tiles. The old code resolved the same path
// twice, so a miss cost two round trips and produced two silent
// failures instead of one honest message.
async function _loadDetailAssets() {
  const ctx = _detailContext;
  const box = document.getElementById('cd-blocks');
  if (!ctx || !box) return;

  const say = msg => {
    if (_detailContext !== ctx) return;
    _detailBlocks = [];
    box.innerHTML = `<p class="prose dim">${escHtml(msg)}</p>`;
  };

  let drive, folder;
  try {
    drive = await resolveDrive(HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary);
    console.info(`[Assets] library “${drive.name}” — looking for ${ctx.folderRoot} ▸ ${ctx.campaignFolder}`);
    folder = await _resolveItemFolder(drive.id, ctx.folderRoot, ctx.campaignFolder);
  } catch (e) {
    console.warn('[Assets] could not reach the document library:', e.message);
    say(`Couldn’t reach the document library — ${e.message}`);
    return;
  }

  if (!folder) {
    say(`Nothing is filed under ${ctx.folderRoot} ▸ ${ctx.campaignFolder} yet.`);
    return;
  }
  if (_detailContext !== ctx) return;

  _detailFolder = { driveId: drive.id, id: folder.id, name: folder.name };
  _loadDetailHero(ctx, drive.id, folder.id);

  // Everything under this campaign or launch, open on the page.
  const panel = document.getElementById('cd-asset-panel');
  if (panel) panel.innerHTML = '';

  const n = await openEverything('cd-blocks', drive.id, folder.id, {
    rootLabel: folder.name,
    emptyText: `“${folder.name}” is empty at the moment. Anything filed into it in SharePoint appears here on its own.`,
  });
  if (_detailContext !== ctx) return;
  console.info(`[Assets] “${folder.name}” — ${n} file(s) opened in the page.`);
}

// The item's own artwork, from its asset folder — same source the cards
// use, so the detail page matches the card you clicked.
async function _loadDetailHero(ctx, driveId, folderId) {
  const el = document.getElementById('dt-hero-img');
  if (!el) return;
  try {
    const url = await folderHeroImage(driveId, folderId);
    if (!url) return;
    if (_detailContext !== ctx) return;          // navigated away meanwhile
    el.style.backgroundImage = `url('${safeCssUrl(url)}')`;
    el.classList.add('has-img');
  } catch (_) { /* the initials placeholder is a fine fallback */ }
}

let _detailBlocks = [];

// The item's own folder, once resolved — so opening a tile is a single
// call on a known id rather than walking the path again by name.
let _detailFolder = null;

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

// Find a child folder for a name. Exact (case-insensitive) first, then
// letters-and-digits-only, then the longest COMMON PREFIX — so the
// "Commander Fire Blankets" campaign finds a folder marketing named
// "Commander fire blanket assets" (plurals and "-assets" suffixes are
// exactly where containment matching falls over). Exact matching alone
// is why detail pages kept falling back to the config tile list.
//
// Digit guard: two names that both carry numbers must carry the SAME
// numbers — "FSE 2026" must never fuzzy-match "FSE 2027".
function _digitsOf(slug) {
  return (String(slug).match(/\d+/g) || []).join(',');
}

async function _findChildFolder(driveId, parentId, name) {
  const target = String(name || '').trim().toLowerCase();
  if (!target) return null;
  const items = (await fetchDriveChildren(driveId, parentId)).filter(x => x.folder);

  let hit = items.find(x => String(x.name).trim().toLowerCase() === target);
  if (hit) return hit;

  const key = _slugKey(target);
  if (!key) return null;
  hit = items.find(x => _slugKey(x.name) === key);
  if (hit) return hit;

  const keyDigits = _digitsOf(key);
  let best = null, bestLen = 0;
  for (const x of items) {
    const k = _slugKey(x.name);
    if (!k) continue;
    const kDigits = _digitsOf(k);
    if (keyDigits && kDigits && keyDigits !== kDigits) continue;

    let p = 0;
    const n = Math.min(k.length, key.length);
    while (p < n && k[p] === key[p]) p++;
    if (p > bestLen) { best = x; bestLen = p; }
  }
  // A real match shares most of the shorter name, not just a word.
  const minLen = best ? Math.min(_slugKey(best.name).length, key.length) : 0;
  return (bestLen >= 5 && bestLen >= minLen * 0.6) ? best : null;
}

// Resolve  Documents/<folderRoot>/<campaignFolder>/<block.folder>  and open
// its file(s) in-hub. One file opens straight into the preview; several are
// listed in a panel; none shows a friendly "not set up yet" note.
async function openDetailAsset(blockIdx) {
  const block = (_detailBlocks || [])[blockIdx];
  const panel = document.getElementById('cd-asset-panel');
  if (!block || !panel) return;

  panel.innerHTML = `<p class="prose dim">Opening &ldquo;${escHtml(block.label)}&rdquo;&hellip;</p>`;

  // The tiles now carry the real drive and item ids, resolved once when
  // the page loaded. No walking the path by name a second time, which is
  // what used to fail here and leave the panel empty.
  try {
    await _showAssetFolder(block.driveId || (_detailFolder && _detailFolder.driveId),
                           block.id, block.label, panel);
  } catch (e) {
    panel.innerHTML = `<p class="prose dim">${escHtml(e.message)}</p>`;
  }
}

// Drill into a sub-folder of an asset folder without leaving the page.
async function openDetailSubfolder(driveId, folderId, label) {
  const panel = document.getElementById('cd-asset-panel');
  if (!panel) return;
  try {
    await _showAssetFolder(driveId, folderId, label, panel);
  } catch (e) {
    panel.innerHTML = `<p class="prose dim">${escHtml(e.message)}</p>`;
  }
}

async function _showAssetFolder(driveId, folderId, label, panel) {
  if (!driveId || !folderId) throw new Error(`“${label}” could not be opened — its folder is missing.`);

  const kids    = await fetchDriveChildren(driveId, folderId);
  const folders = kids.filter(x => x.folder).sort((a, b) => String(a.name).localeCompare(String(b.name)));
  const files   = kids.filter(x => !x.folder).sort((a, b) => String(a.name).localeCompare(String(b.name)));

  if (!folders.length && !files.length) {
    panel.innerHTML = `<h3 class="dt-panel-head">${escHtml(label)}<span>0</span></h3>
      <p class="prose dim">Nothing in &ldquo;${escHtml(label)}&rdquo; yet.</p>`;
    return;
  }

  _lastAssetFiles = files;

  const folderRows = folders.map(k => {
    const n = (k.folder && k.folder.childCount) || 0;
    const args = `'${escAttr(driveId)}','${escAttr(k.id)}','${escAttr(String(k.name).replace(/'/g, ''))}'`;
    return `
      <div class="lib-file folder" role="button" tabindex="0"
           onclick="openDetailSubfolder(${args})"
           onkeydown="if(event.key==='Enter')openDetailSubfolder(${args})">
        <span class="lib-file-ico folder">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="17" height="17"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
        </span>
        <span class="lib-file-main">
          <span class="lib-file-name">${escHtml(k.name)}</span>
          <span class="lib-file-meta">${n} item${n === 1 ? '' : 's'}</span>
        </span>
        <span class="lib-file-go">Open →</span>
      </div>`;
  }).join('');

  // Always LIST the files rather than auto-opening a single one — people
  // want the download and copy-link buttons at least as often as they
  // want to read it.
  panel.innerHTML = `
    <h3 class="dt-panel-head">${escHtml(label)}<span>${files.length}</span></h3>
    <div class="lib-files">${folderRows}${files.map(f => libFileRow(f)).join('')}</div>`;

  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function _eventYear(name) {
  const m = String(name || '').match(/(19|20)\d{2}/);
  return m ? parseInt(m[0], 10) : null;
}

// ═══ Trade, events & training ════════════════════════════════
//
// REBUILT 26 Aug 2026 (second round) into the same editorial shape as
// Launches and Campaigns, and given the training half David asked for:
// "I think we should add more on the training events and things like
// that in there."
//
// Events still come from the folders in Documents ▸ Events — one
// folder per event, everything the sales team needs inside it.
// Training comes from the Training Events list, and staff can now book
// themselves onto a session in one click.

// 1 SEP 2026 — marketing: "Can this page please go as it was. Divided
// into: Exhibitions / Customer events / Training. Don't want that to
// feel it's only for FSE."
//
// It felt like an FSE page because it WAS one: Documents ▸ Events held
// FSE 2025, FSE 2026 and FSE 2027 and nothing else, and the biggest
// thing on the screen was a full-width spread for whichever of them
// came next. Both are fixed here. The spread is gone — no single event
// gets to be the page any more — and the events are grouped into the
// three sections marketing asked for.
//
// The grouping lives in SharePoint (see HUB_CONFIG.tradeEvents): a
// folder per category under Events, with the event folders inside it.
// Anything still loose at the top level keeps showing under the
// fallback category, so the page never empties out while the folders
// are being moved.
function _evCatOf(name) {
  const cats = (HUB_CONFIG.tradeEvents && HUB_CONFIG.tradeEvents.categories) || [];
  const key = _slugKey(name);
  for (const c of cats) {
    const names = [c.folder].concat(c.aliases || []);
    if (names.some(x => _slugKey(x) === key)) return c;
  }
  return null;
}

async function loadTradeEvents() {
  const host = document.getElementById('ev-index');
  if (!host) return;

  try {
    const drive  = await resolveDrive(HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary);
    const rootNm = (HUB_CONFIG.tradeEvents && HUB_CONFIG.tradeEvents.folder) || 'Events';
    const root   = await _findChildFolder(drive.id, null, rootNm);
    if (!root) throw new Error(`No "${rootNm}" folder in the document library yet.`);

    const cats     = (HUB_CONFIG.tradeEvents && HUB_CONFIG.tradeEvents.categories) || [];
    const fallback = (HUB_CONFIG.tradeEvents && HUB_CONFIG.tradeEvents.fallback)
                     || (cats[0] && cats[0].key) || 'exhibitions';
    const thisYear = new Date().getFullYear();

    const top = (await fetchDriveChildren(drive.id, root.id)).filter(x => x.folder);

    // A top-level folder is either a category container or a loose event.
    const found = [];
    const loose = [];
    for (const t of top) {
      const c = _evCatOf(t.name);
      if (c) found.push({ cat: c.key, folder: t });
      else   loose.push(t);
    }
    loose.forEach(t => found.push({ cat: fallback, folder: t, direct: true }));

    // Read the category containers one level down, in parallel.
    const groups = [];
    await Promise.all(found.map(async row => {
      if (row.direct) { groups.push({ cat: row.cat, item: row.folder }); return; }
      let kids = [];
      try { kids = (await fetchDriveChildren(drive.id, row.folder.id)).filter(x => x.folder); }
      catch (e) { console.info('[Events] could not read "' + row.folder.name + '":', e.message); }
      kids.forEach(k => groups.push({ cat: row.cat, item: k }));
    }));

    _eventFolders = groups.map(g => {
      const k = g.item;
      const year = _eventYear(k.name);
      return {
        id: k.id, name: k.name, year: year,
        count: (k.folder && k.folder.childCount) || 0,
        modified: k.lastModifiedDateTime,
        driveId: drive.id, cat: g.cat,
        // No year in the name? Treat it as current/ongoing.
        upcoming: year === null || year >= thisYear,
      };
    });

    const order = (a, b) =>
      (a.upcoming === b.upcoming ? 0 : a.upcoming ? -1 : 1) ||
      (a.upcoming ? (a.year || thisYear) - (b.year || thisYear)
                  : (b.year || 0) - (a.year || 0)) ||
      String(a.name).localeCompare(String(b.name));

    const evCard = (e) => {
      const i = _eventFolders.indexOf(e);
      return `
      <article class="px-card" data-tone="${e.upcoming ? 'green' : 'grey'}" style="--i:${i}"
               role="button" tabindex="0"
               onclick="openEventFolder(${i})" onkeydown="if(event.key==='Enter')openEventFolder(${i})">
        <div class="px-card-media" id="px-img-event-${i}">
          <span class="px-card-initials">${escHtml(_pxInitials(e.name))}</span>
          <span class="px-badge"><span class="px-badge-dot ${e.upcoming ? 'green' : 'grey'}"></span>${e.upcoming ? 'Upcoming' : 'Previous'}</span>
        </div>
        <div class="px-card-body">
          <div class="px-card-eyebrow">${escHtml(e.year ? String(e.year) : 'Ongoing')}</div>
          <h3 class="px-card-title">${escHtml(e.name)}</h3>
          <div class="px-card-meta">${e.count} item${e.count === 1 ? '' : 's'} in the pack</div>
        </div>
        <span class="px-card-go">Open <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" width="13" height="13"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg></span>
      </article>`;
    };

    const section = (c) => {
      const mine = _eventFolders.filter(e => e.cat === c.key).sort(order);
      const next = mine.find(e => e.upcoming);
      return `
      <section class="ev-sec">
        <div class="ev-sec-head">
          <div>
            <h2 class="ev-sec-title">${escHtml(c.label)}</h2>
            ${c.sub ? `<p class="ev-sec-sub">${escHtml(c.sub)}</p>` : ''}
          </div>
          <span class="ev-sec-n">${mine.length
            ? `${mine.length} ${mine.length === 1 ? 'event' : 'events'}${next ? ' · next: ' + escHtml(next.name) : ''}`
            : 'Nothing yet'}</span>
        </div>
        ${mine.length
          ? `<div class="px-grid" data-filter="all">${mine.map(evCard).join('')}</div>`
          : `<p class="ev-sec-empty">Nothing filed here yet. Add a folder for each one under
             <strong>Documents ▸ ${escHtml(rootNm)} ▸ ${escHtml(c.folder)}</strong> on the MarketingHub
             SharePoint site and it appears here, with everything inside it.</p>`}
      </section>`;
    };

    host.innerHTML = `
      <div class="ev-wrap">
        ${cats.map(section).join('')}
      </div>
      <section id="ev-training"></section>`;

    // Event artwork, painted in once SharePoint answers.
    Promise.all(_eventFolders.slice(0, 16).map(async e => {
      const url = await folderHeroImage(e.driveId, e.id);
      if (!url) return;
      const el = document.getElementById('px-img-event-' + _eventFolders.indexOf(e));
      if (el) { el.style.backgroundImage = `url('${safeCssUrl(url)}')`; el.classList.add('has-img'); }
    })).catch(() => {});

    renderTrainingBand();

  } catch (e) {
    const msg = e.message === 'NOT_FOUND'
      ? 'Could not reach the document library — check you have access to the MarketingHub site.'
      : e.message;
    host.innerHTML = `<p class="sp-error">${escHtml(msg)}</p>`;
  }
}

function filterEvents(tone, btn) {
  const grid = document.getElementById('px-event-grid');
  if (grid) grid.setAttribute('data-filter', tone);
  if (btn && btn.parentElement) {
    btn.parentElement.querySelectorAll('.px-chip').forEach(b => b.classList.toggle('active', b === btn));
  }
}

// ── Training sessions, with one-click sign-up ─────────────────
// Marketing put a session in the "Training Events" list; staff book
// themselves onto it from here. The booking is written to the
// "Training Signups" list so marketing can see who's coming, and the
// same click drops the session into the person's own calendar.

const TRAIN = { sessions: [], signups: [], busy: null, loaded: false };

// auth.js stores the signed-in user as { displayName, mail } — NOT the
// raw MSAL account. Reading .username here is why "Book me on" told a
// signed-in David to sign in first. Same shape the poll reads.
function _trainMyEmail() {
  const a = (window.AUTH && window.AUTH.account) || {};
  return String(a.mail || a.username || a.email || '').toLowerCase();
}

function _trainMyName() {
  const a = (window.AUTH && window.AUTH.account) || {};
  return a.displayName || a.name || a.mail || '';
}

function _trainSignupList() {
  return ((HUB_CONFIG.trainingSignup || {}).list) || 'Training Signups';
}

function _trainMine(session) {
  const me = _trainMyEmail().toLowerCase();
  if (!me) return null;
  return TRAIN.signups.find(r =>
    String(r.fields.SessionId || '') === String(session.id) &&
    String(r.fields.Attendee || '').toLowerCase() === me) || null;
}

function _trainGoing(session) {
  return TRAIN.signups.filter(r => String(r.fields.SessionId || '') === String(session.id)).length;
}

async function renderTrainingBand() {
  const host = document.getElementById('ev-training');
  if (!host) return;

  host.innerHTML = `<div class="tr-band"><div class="tr-inner">
    <div class="skeleton sk-line med"></div><div class="skeleton sk-line"></div></div></div>`;

  try {
    const rows = await _fetchListRows(HUB_CONFIG.training.list, 100);
    const now = new Date(); now.setHours(0, 0, 0, 0);

    TRAIN.sessions = rows
      .map(r => Object.assign({ id: r.id }, r.fields || {}))
      .filter(f => f.TrainingDate && !isNaN(new Date(f.TrainingDate)))
      .filter(f => new Date(f.EndDate || f.TrainingDate) >= now)
      .sort((a, b) => String(a.TrainingDate).localeCompare(String(b.TrainingDate)));

    // Who's already booked. Read-only, so it never asks for consent.
    try {
      TRAIN.signups = await _fetchListRows(_trainSignupList(), 500);
    } catch (_) {
      TRAIN.signups = [];   // list not created yet — sign-up explains
    }
    TRAIN.loaded = true;
  } catch (e) {
    host.innerHTML = `<div class="tr-band"><div class="tr-inner">
      <p class="sp-error">Couldn't load training sessions: ${escHtml(e.message)}</p></div></div>`;
    return;
  }

  renderTrainingList();
}

function renderTrainingList() {
  const host = document.getElementById('ev-training');
  if (!host) return;

  const head = `
    <div class="tr-head">
      <div>
        <h2 class="tr-title">Training &amp; sessions</h2>
        <p class="tr-sub">Internal sessions and external courses. Book yourself on and it lands in your calendar.</p>
      </div>
      <div class="tr-sub-wrap">
        <button class="tr-sub-btn" onclick="subscribeTrainingDates(this)">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><path d="M4 11a9 9 0 0 1 9 9"/><path d="M4 4a16 16 0 0 1 16 16"/><circle cx="5" cy="19" r="1"/></svg>
          Add all dates to my calendar
        </button>
        <span class="tr-sub-note" id="tr-sub-note"></span>
      </div>
    </div>`;

  if (!TRAIN.sessions.length) {
    host.innerHTML = `<div class="tr-band"><div class="tr-inner">${head}
      <div class="tr-empty">Nothing scheduled at the moment. Sessions marketing add to the
      <strong>Training Events</strong> list appear here, and on the calendar on the home page.</div>
    </div></div>`;
    return;
  }

  host.innerHTML = `<div class="tr-band"><div class="tr-inner">${head}
    <div class="tr-list">${TRAIN.sessions.map((s, i) => _trainCard(s, i)).join('')}</div>
  </div></div>`;
}

function _trainCard(s, i) {
  const d = new Date(s.TrainingDate);
  const day = String(d.getDate());
  const mon = d.toLocaleDateString('en-GB', { month: 'short' });
  const kind = /external/i.test(s.TrainingType || '') ? 'external' : 'internal';
  const mine = _trainMine(s);
  const going = _trainGoing(s);
  const busy = TRAIN.busy === s.id;

  const meta = [
    s.Trainer, s.Location,
    s.EndDate && s.EndDate !== s.TrainingDate ? 'until ' + fmtSpDate(s.EndDate) : '',
  ].filter(Boolean).map(escHtml).join(' · ');

  const button = busy
    ? `<button class="tr-book working" disabled>Saving…</button>`
    : mine
      ? `<button class="tr-book booked" onclick="cancelTraining('${escAttr(s.id)}')">
           <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" width="13" height="13"><polyline points="20 6 9 17 4 12"/></svg>
           You're booked on
         </button>`
      : `<button class="tr-book" onclick="signUpTraining('${escAttr(s.id)}')">Book me on</button>`;

  return `
    <article class="tr-card" style="--i:${i}">
      <div class="tr-date ${kind}">
        <span class="d">${escHtml(day)}</span>
        <span class="m">${escHtml(mon)}</span>
      </div>
      <div class="tr-main">
        <span class="tr-tag ${kind}">${kind === 'external' ? 'External' : 'Internal'}</span>
        <h3 class="tr-name">${escHtml(s.Title || 'Training session')}</h3>
        ${meta ? `<div class="tr-meta">${meta}</div>` : ''}
        ${s.Notes ? `<p class="tr-notes">${escHtml(s.Notes)}</p>` : ''}
      </div>
      <div class="tr-act">
        ${button}
        <span class="tr-going">${going ? `${going} booked on` : 'Be the first'}</span>
        ${s.Link ? `<a class="tr-link" href="${escAttr(safeUrl(linkOf(s.Link), '#'))}" target="_blank" rel="noopener">Joining details →</a>` : ''}
        ${mine ? `<button class="tr-link" onclick="addTrainingToCalendar('${escAttr(s.id)}')">Add to calendar again</button>` : ''}
      </div>
    </article>`;
}

function _trainSession(id) {
  return TRAIN.sessions.find(s => String(s.id) === String(id));
}

// Every session on this page as one .ics. Downloading a calendar file
// is only half a feature if nobody says what to do with it — the
// button explains itself inline after the click, and stays explained.
function subscribeTrainingDates(btn) {
  if (!TRAIN.sessions.length) { showToast('Nothing scheduled to add yet'); return; }

  const events = TRAIN.sessions.map(s => [
    'BEGIN:VEVENT',
    `UID:cf-training-${_icsDate(s.TrainingDate)}-${String(s.id)}@checkfire-hub`,
    `DTSTAMP:${_icsDate(new Date().toISOString())}T090000Z`,
    `DTSTART;VALUE=DATE:${_icsDate(s.TrainingDate)}`,
    `DTEND;VALUE=DATE:${_icsNextDay(s.EndDate || s.TrainingDate)}`,
    `SUMMARY:${_icsEscape(s.Title || 'CheckFire training')}`,
    s.Location ? `LOCATION:${_icsEscape(s.Location)}` : '',
    s.Trainer ? `DESCRIPTION:${_icsEscape('Trainer: ' + s.Trainer)}` : '',
    'END:VEVENT',
  ].filter(Boolean).join('\r\n'));

  const ics = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//CheckFire//Marketing Hub//EN',
    'CALSCALE:GREGORIAN', 'METHOD:PUBLISH', ...events, 'END:VCALENDAR'].join('\r\n');

  const blob = new Blob([ics], { type: 'text/calendar;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'checkfire-training-dates.ics';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);

  const note = document.getElementById('tr-sub-note');
  if (note) note.innerHTML =
    `<strong>checkfire-training-dates.ics</strong> is in your downloads — open it and Outlook adds all ${TRAIN.sessions.length} date${TRAIN.sessions.length === 1 ? '' : 's'}.`;
  if (btn) btn.blur();
}

// One .ics for one session — this is what actually puts it in Outlook.
function addTrainingToCalendar(id) {
  const s = _trainSession(id);
  if (!s) return;
  const start = _icsDate(s.TrainingDate);
  const end   = _icsNextDay(s.EndDate || s.TrainingDate);
  const desc  = [s.Trainer ? 'Trainer: ' + s.Trainer : '', s.Notes || ''].filter(Boolean).join('\\n');

  const ics = [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//CheckFire//Marketing Hub//EN',
    'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    `UID:cf-training-${_icsDate(s.TrainingDate)}-${String(s.id)}@checkfire-hub`,
    `DTSTAMP:${_icsDate(new Date().toISOString())}T090000Z`,
    `DTSTART;VALUE=DATE:${start}`,
    `DTEND;VALUE=DATE:${end}`,
    `SUMMARY:${_icsEscape(s.Title || 'CheckFire training')}`,
    s.Location ? `LOCATION:${_icsEscape(s.Location)}` : '',
    desc ? `DESCRIPTION:${_icsEscape(desc)}` : '',
    'END:VEVENT', 'END:VCALENDAR',
  ].filter(Boolean).join('\r\n');

  const blob = new Blob([ics], { type: 'text/calendar' });
  const url  = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = (s.Title || 'training').replace(/[^A-Za-z0-9 _-]/g, '').trim().slice(0, 60) + '.ics';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

async function signUpTraining(id) {
  const s = _trainSession(id);
  if (!s || TRAIN.busy) return;
  const me = _trainMyEmail();
  if (!me) { showToast('Sign in first'); return; }

  TRAIN.busy = s.id;
  renderTrainingList();

  try {
    const token = (typeof getWriteToken === 'function') ? await getWriteToken() : null;
    if (!token) throw new Error('NO_WRITE');

    const siteId = await getSiteId();
    const list   = _trainSignupList();
    const url    = `${GRAPH_BASE}/sites/${siteId}/lists/${encodeURIComponent(list)}/items`;

    const base = {
      Title:     String(s.Title || 'Training').slice(0, 255),
      SessionId: String(s.id),
      Attendee:  me,
    };

    const post = fields => fetch(url, {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields }),
    });

    // Optional columns first; SharePoint rejects the whole write if a
    // column is missing, so fall back to the three that must exist.
    let res = await post(Object.assign({}, base, {
      AttendeeName: _trainMyName(),
      SessionDate:  s.TrainingDate,
    }));
    if (res.status === 400) res = await post(base);

    if (res.status === 404) throw new Error('NO_LIST');
    if (!res.ok) throw new Error('SharePoint returned ' + res.status);

    const created = await res.json();
    TRAIN.signups.push({ id: created.id, fields: created.fields || base });

    addTrainingToCalendar(id);
    showToast('Booked on — it’s in your calendar');
  } catch (e) {
    if (e.message === 'NO_WRITE') {
      showToast('Needs permission to save your booking — ask IT to approve hub write access');
    } else if (e.message === 'NO_LIST') {
      showToast(`The "${_trainSignupList()}" list hasn't been created yet`);
      addTrainingToCalendar(id);   // at least get it in their calendar
    } else {
      showToast('Not saved — ' + e.message);
    }
  }

  TRAIN.busy = null;
  renderTrainingList();
}

async function cancelTraining(id) {
  const s = _trainSession(id);
  if (!s || TRAIN.busy) return;
  const mine = _trainMine(s);
  if (!mine) return;

  TRAIN.busy = s.id;
  renderTrainingList();

  try {
    const token = (typeof getWriteToken === 'function') ? await getWriteToken() : null;
    if (!token) throw new Error('NO_WRITE');
    const siteId = await getSiteId();
    const list   = _trainSignupList();
    const res = await fetch(
      `${GRAPH_BASE}/sites/${siteId}/lists/${encodeURIComponent(list)}/items/${encodeURIComponent(mine.id)}`,
      { method: 'DELETE', headers: { Authorization: 'Bearer ' + token } });
    if (!res.ok && res.status !== 204) throw new Error('SharePoint returned ' + res.status);
    TRAIN.signups = TRAIN.signups.filter(r => r !== mine);
    showToast('Booking cancelled');
  } catch (e) {
    showToast(e.message === 'NO_WRITE'
      ? 'Needs permission to change your booking'
      : 'Not cancelled — ' + e.message);
  }

  TRAIN.busy = null;
  renderTrainingList();
}

// ── Event pack (the drill-down) ───────────────────────────────

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

  // 1 Sep 2026 (round 2) — "open everything … like a website". An
  // event pack is the same problem as a campaign: you wanted the stand
  // plans and the artwork, not a folder tree to walk. The crumbs bar is
  // now a jump list built by openEverything instead.
  const crumbs = document.getElementById('ev-crumbs');
  if (crumbs) crumbs.innerHTML = '';
  openEverything('ev-documents-grid', ev.driveId, ev.id, {
    rootLabel: ev.name,
    emptyText: `Nothing has been filed into “${ev.name}” yet. Add it under Documents ▸ Events in SharePoint and it appears here.`,
  });
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
// A live poll card on the home page, driven by the SharePoint "Polls"
// list. It reads the list marketing actually built:
//
//   Title       short code, e.g. "NPD" / "CO2"
//   Question    the question people answer   (falls back to Title)
//   Status      choice — only "Open" is shown (Draft/Closed are hidden)
//   OpensDate   optional — hidden before this day
//   ClosesDate  optional — hidden after this day
//   PollURL     optional — shows an "Open the full form" button
//   Options     OPTIONAL multi-line, one choice per line
//
// Two modes, chosen automatically:
//   * Options filled in (2+ lines)  → click-to-vote, results as % bars
//   * Options empty                 → free-text answer box in the card
//
// Either way the answer is written to the "Poll Votes" list:
//   Title (the answer, trimmed to fit) · PollId · Voter (email)
//   Answer (multi-line, full text) — used if the column exists
//
// Reading uses the hub's normal read-only token. Answering asks for
// Sites.ReadWrite.All the first time — see getWriteToken() in auth.js.

const POLL = { poll: null, votes: [], busy: false };

// Like fetchListItems() but keeps the SharePoint item id, which the
// poll needs to tie answers to a question. Deliberately uncached so a
// new answer shows up straight away.
async function _fetchListRows(listName, top) {
  const siteId = await getSiteId();
  const data = await graphFetch(
    `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items?expand=fields&$top=${top || 100}`
  );
  return (data.value || []).map(i => ({ id: i.id, fields: i.fields || {} }));
}

// SharePoint date columns come back as UTC instants, and the MarketingHub
// site is still on Pacific time, so a "13 Aug" date arrives as
// 2026-08-13T07:00:00Z. Compare the calendar day only — never the
// instant — or a poll opening today looks like it opens in the future.
function _pollDay(v) {
  if (!v) return '';
  const m = String(v).match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  const d = new Date(v);
  if (isNaN(d)) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function _todayDay() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function _pollQuestion(f) {
  return String(f.Question || f.Title || 'Quick poll').trim();
}

// "Options" is optional. One option per line. If the column was created
// as rich text it arrives as HTML, so tags become line breaks.
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
  // Status column (marketing's shape): only "Open" runs.
  const status = String(f.Status || '').trim().toLowerCase();
  if (status && status !== 'open') return false;

  // Active Yes/No (the original shape) still honoured.
  const a = f.Active;
  if (a === false || a === 0 || a === 'No' || a === 'false') return false;

  const today  = _todayDay();
  const opens  = _pollDay(f.OpensDate  || f.StartDate);
  const closes = _pollDay(f.ClosesDate || f.EndDate);
  if (opens  && opens  > today) return false;
  if (closes && closes < today) return false;
  return true;
}

function _pollMyEmail() {
  return String((window.AUTH && window.AUTH.account && window.AUTH.account.mail) || '').toLowerCase();
}

function _pollNote(msg) {
  const el = document.getElementById('poll-note');
  if (el) el.textContent = msg || '';
}

function _pollVotesListName() {
  return (HUB_CONFIG.polls && HUB_CONFIG.polls.votesList) || 'Poll Votes';
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
    const live = rows.filter(r => _pollIsLive(r.fields));
    // Newest first — SharePoint item ids increase.
    const poll = live.sort((a, b) => Number(b.id) - Number(a.id))[0];

    if (!poll) {
      POLL.poll = null;
      body.innerHTML =
        '<p class="poll-empty">No poll running right now.<br>' +
        '<span class="poll-hint">Set a question to <strong>Open</strong> in the Polls list to start one.</span></p>';
      return;
    }

    POLL.poll = poll;
    try {
      const votes = await _fetchListRows(_pollVotesListName(), 500);
      POLL.votes = votes.filter(v => String(v.fields.PollId || '') === String(poll.id));
    } catch (e) {
      POLL.votes = [];
      console.info('[Polls] votes list not readable yet:', e.message);
    }
    renderPoll();
  } catch (e) {
    console.info('[Polls] not loaded:', e.message);
    body.innerHTML = '<p class="poll-empty">Poll unavailable.<br><span class="poll-hint">Check the <strong>Polls</strong> list on the MarketingHub site.</span></p>';
  }
}

function _pollAnswerText(v) {
  return String(v.fields.Answer || v.fields.Title || '');
}

function renderPoll() {
  const body = document.getElementById('home-poll-body');
  if (!body || !POLL.poll) return;

  const f     = POLL.poll.fields;
  const opts  = _pollOptions(f);
  const me    = _pollMyEmail();
  const mine  = POLL.votes.find(v => String(v.fields.Voter || '').toLowerCase() === me);
  const total = POLL.votes.length;
  // safeUrl() strips anything that isn't http(s) — a hyperlink column
  // is staff-entered, but it still goes through the same guard as every
  // other URL in the hub.
  const url   = safeUrl(linkOf(f.PollURL), '');

  const question = escHtml(_pollQuestion(f));
  const choice   = opts.length >= 2;

  let inner;

  if (mine && choice) {
    // Answered, multiple choice — show the split.
    const picked = _pollAnswerText(mine);
    const counts = {};
    opts.forEach(o => { counts[o] = 0; });
    POLL.votes.forEach(v => {
      const t = _pollAnswerText(v);
      if (counts[t] !== undefined) counts[t]++;
    });
    inner = opts.map(o => {
      const n   = counts[o] || 0;
      const pct = total ? Math.round((n / total) * 100) : 0;
      return `
        <div class="poll-res${o === picked ? ' mine' : ''}">
          <div class="poll-res-top">
            <span class="poll-res-lbl">${escHtml(o)}${o === picked ? ' <span class="poll-tick">✓</span>' : ''}</span>
            <span class="poll-res-pct">${pct}%</span>
          </div>
          <div class="poll-bar"><span style="width:${pct}%"></span></div>
        </div>`;
    }).join('');

  } else if (mine) {
    // Answered, free text — show it back. Other people's answers stay in
    // SharePoint rather than being splashed across the home page.
    inner = `
      <div class="poll-answered">
        <div class="poll-answered-lbl">Your answer</div>
        <div class="poll-answered-text">${escHtml(_pollAnswerText(mine))}</div>
      </div>`;

  } else if (choice) {
    inner = opts.map((o, i) => `
      <button type="button" class="poll-opt" onclick="votePoll(${i})">${escHtml(o)}</button>
    `).join('');

  } else {
    inner = `
      <textarea id="poll-text" class="poll-text" rows="3" maxlength="1000"
                placeholder="Type your answer…"
                onkeydown="if(event.key==='Enter'&&(event.metaKey||event.ctrlKey))submitPollAnswer()"></textarea>
      <button type="button" class="poll-send" onclick="submitPollAnswer()">Send answer</button>`;
  }

  const link = url
    ? `<a class="poll-link" href="${escAttr(url)}" target="_blank" rel="noopener">Open the full form →</a>`
    : '';

  const countLbl = choice
    ? `${total} vote${total === 1 ? '' : 's'}`
    : `${total} answer${total === 1 ? '' : 's'}`;

  body.innerHTML = `
    <div class="poll-q">${question}</div>
    <div class="poll-opts">${inner}</div>
    ${link}
    <div class="poll-foot">
      <span>${countLbl}${mine ? ' · thanks' : ''}</span>
      <span id="poll-note" class="poll-note"></span>
    </div>`;
}

// Shared write path for both modes.
async function _pollSubmit(answer) {
  if (POLL.busy || !POLL.poll) return;
  const text = String(answer || '').trim();
  if (!text) { _pollNote('Type something first.'); return; }

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
    const listName = _pollVotesListName();
    const url      = `${GRAPH_BASE}/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`;

    const base = {
      Title:  text.slice(0, 255),
      PollId: String(POLL.poll.id),
      Voter:  _pollMyEmail(),
    };

    const post = fields => fetch(url, {
      method:  'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body:    JSON.stringify({ fields }),
    });

    // Try with the full-text Answer column; if that column doesn't exist
    // SharePoint rejects the whole write, so fall back to Title only.
    let res = await post({ ...base, Answer: text });
    if (res.status === 400) res = await post(base);

    if (res.status === 404) throw new Error('the "' + listName + '" list does not exist yet');
    if (!res.ok) throw new Error('SharePoint returned ' + res.status);

    const created = await res.json();
    POLL.votes.push({ id: created.id, fields: created.fields || { ...base, Answer: text } });
    renderPoll();
  } catch (e) {
    _pollNote('Not saved — ' + e.message);
  }
  POLL.busy = false;
}

function votePoll(i) {
  if (!POLL.poll) return;
  const opts = _pollOptions(POLL.poll.fields);
  if (!opts[i]) return;
  _pollSubmit(opts[i]);
}

function submitPollAnswer() {
  const el = document.getElementById('poll-text');
  if (!el) return;
  _pollSubmit(el.value);
}


// ═══ Site search ═════════════════════════════════════════════
//
// 26 Aug 2026 (fix 4). The header box used to open Ember and ask her the
// question. David: "I'd rather the search bar be able to search the site
// not talk to ember through it." So Enter now runs a real search over the
// hub's own content and lands on a results PAGE. Ember is still there on
// her own button for the questions that need reading and reasoning; this
// is for the far more common "where is that thing".
//
// Two halves, run in parallel:
//   ITEMS  — launches, campaigns, events and training sessions, matched
//            against the fields people actually search by.
//   FILES  — a Graph drive search across the Marketing Library and the
//            Product Portal, so a certificate is findable by name from
//            anywhere in the hub.
// Files open in the reader page; items land on their own page with the
// detail already open.

const SRCH = { q: '', seq: 0 };

function _srchWords(q) {
  return String(q || '').toLowerCase().split(/\s+/).map(w => w.trim()).filter(Boolean);
}

function _srchHit(hay, words) {
  const h = String(hay || '').toLowerCase();
  return words.every(w => h.includes(w));
}

async function openSiteSearch(q) {
  const query = String(q || '').trim();
  if (!query) return;

  const host = document.getElementById('srch-results');
  if (!host) return;

  const seq = ++SRCH.seq;
  SRCH.q = query;

  const title = document.getElementById('srch-title');
  const count = document.getElementById('srch-count');
  if (title) title.innerHTML = 'Results for <em>' + escHtml(query) + '</em>';
  if (count) count.textContent = '';

  if (typeof toggleEmber === 'function') toggleEmber(false);
  if (typeof showPage === 'function') showPage('search');
  if (typeof updateNavActive === 'function') updateNavActive('search');

  host.innerHTML = '<div class="srch-wait"><span class="rdr-spin"></span>Searching the hub…</div>';

  const words = _srchWords(query);
  const [items, files] = await Promise.all([
    _srchItems(words).catch(e => { console.warn('[Search] items:', e.message); return []; }),
    _srchFiles(query).catch(e => { console.warn('[Search] files:', e.message); return []; }),
  ]);

  if (seq !== SRCH.seq) return;    // a newer search overtook this one

  const total = items.length + files.length;
  if (count) {
    count.textContent = total
      ? `${total} result${total === 1 ? '' : 's'} — ${items.length} page${items.length === 1 ? '' : 's'}, ${files.length} document${files.length === 1 ? '' : 's'}`
      : '';
  }

  if (!total) {
    host.innerHTML = `<div class="srch-empty">
      Nothing in the hub matches &ldquo;${escHtml(query)}&rdquo;.<br>
      Try fewer words, or ask Ember — she reads the documents themselves.
    </div>`;
    return;
  }

  let html = '';
  if (items.length) {
    html += `<div class="srch-group">
      <div class="srch-group-head">Pages &amp; items <span>${items.length}</span></div>
      <div class="srch-items">${items.map(_srchItemHtml).join('')}</div>
    </div>`;
  }
  if (files.length) {
    html += `<div class="srch-group">
      <div class="srch-group-head">Documents <span>${files.length}</span></div>
      <div class="lib-files">${files.map(f => libFileRow(f, f._site)).join('')}</div>
    </div>`;
  }
  host.innerHTML = html;
}

function _srchItemHtml(it) {
  const args = `'${escAttr(it.kind)}','${escAttr(String(it.title).replace(/'/g, ''))}'`;
  return `
    <button class="srch-item" onclick="srchOpenItem(${args})">
      <span class="srch-kind ${escAttr(it.kind)}">${escHtml(it.kindLabel)}</span>
      <span class="srch-item-main">
        <span class="srch-item-name">${escHtml(it.title)}</span>
        ${it.meta ? `<span class="srch-item-meta">${escHtml(it.meta)}</span>` : ''}
      </span>
    </button>`;
}

async function _srchItems(words) {
  const L = HUB_CONFIG.lists || {};
  const [launches, campaigns, training, events] = await Promise.all([
    fetchListItems(L.launches).catch(() => []),
    fetchListItems(L.campaigns).catch(() => []),
    _fetchListRows((HUB_CONFIG.training && HUB_CONFIG.training.list) || 'Training Events', 100)
      .then(rows => rows.map(r => Object.assign({ id: r.id }, r.fields || {})))
      .catch(() => []),
    _srchEventFolders().catch(() => []),
  ]);

  const out = [];

  launches.forEach(f => {
    const hay = [f.Title, f.Description, f.Summary, f.Status, f.ProductCodes, f.Region].join(' ');
    if (_srchHit(hay, words)) {
      out.push({ kind: 'launch', kindLabel: 'Launch', title: f.Title || 'Untitled',
                 meta: [f.Status, fmtSpDate(f.LaunchDate || f.StartDate)].filter(Boolean).join(' · ') });
    }
  });

  campaigns.forEach(f => {
    const hay = [f.Title, f.Description, f.Summary, f.CampaignType, f.Channels, f.Status, f.Region].join(' ');
    if (_srchHit(hay, words)) {
      out.push({ kind: 'campaign', kindLabel: 'Campaign', title: f.Title || 'Untitled',
                 meta: [f.CampaignType, f.Status, fmtSpDate(f.StartDate)].filter(Boolean).join(' · ') });
    }
  });

  events.forEach(e => {
    if (_srchHit(e.name, words)) {
      out.push({ kind: 'event', kindLabel: 'Event', title: e.name,
                 meta: `${e.count} file${e.count === 1 ? '' : 's'}` });
    }
  });

  training.forEach(s => {
    const hay = [s.Title, s.Description, s.Trainer, s.Location, s.Audience].join(' ');
    if (_srchHit(hay, words)) {
      out.push({ kind: 'training', kindLabel: 'Training', title: s.Title || 'Training session',
                 meta: [fmtSpDate(s.TrainingDate), s.Location].filter(Boolean).join(' · ') });
    }
  });

  return out.slice(0, 40);
}

// Events live as folders, not a list. Reuse what the Trade page already
// loaded when it has been visited; otherwise read the folder once.
async function _srchEventFolders() {
  if (_eventFolders && _eventFolders.length) return _eventFolders;
  const drive  = await resolveDrive(HUB_CONFIG.sharepointSite, HUB_CONFIG.documentsLibrary);
  const rootNm = (HUB_CONFIG.tradeEvents && HUB_CONFIG.tradeEvents.folder) || 'Events';
  const root   = await _findChildFolder(drive.id, null, rootNm);
  if (!root) return [];
  // 1 Sep 2026 — events now sit one level down, inside a category
  // folder (Exhibitions / Customer Events). Search walks into those so
  // "FSE 2026" is still findable from the header search box; a folder
  // that isn't a category is still an event in its own right.
  const top = (await fetchDriveChildren(drive.id, root.id)).filter(x => x.folder);
  const out = [];
  const row = k => ({ id: k.id, name: k.name, driveId: drive.id,
                      count: (k.folder && k.folder.childCount) || 0 });
  await Promise.all(top.map(async t => {
    if (!_evCatOf(t.name)) { out.push(row(t)); return; }
    try {
      (await fetchDriveChildren(drive.id, t.id)).filter(x => x.folder).forEach(k => out.push(row(k)));
    } catch (_) { /* a category we can't read simply contributes nothing */ }
  }));
  return out;
}

// Graph's drive search, across both libraries. `id` and `_driveId` are
// selected deliberately — without them a result can't be opened in the
// reader and would fall back to a SharePoint tab.
async function _srchFiles(query) {
  const sites = [
    { label: 'Marketing Library', url: HUB_CONFIG.sharepointSite,    lib: HUB_CONFIG.documentsLibrary },
    { label: 'Product Portal',    url: HUB_CONFIG.productPortalSite, lib: HUB_CONFIG.documentsLibrary },
  ].filter(s => s.url);

  const q = encodeURIComponent(String(query).slice(0, 80).replace(/'/g, "''"));

  const per = await Promise.all(sites.map(async s => {
    try {
      const drive = await resolveDrive(s.url, s.lib);
      const res = await graphFetch(
        `/drives/${drive.id}/root/search(q='${q}')` +
        `?$select=id,name,size,lastModifiedDateTime,webUrl,file,folder&$top=25`);
      return (res.value || [])
        .filter(x => !x.folder)
        .map(x => Object.assign({}, x, { _driveId: drive.id, _site: s.label }));
    } catch (e) {
      console.warn(`[Search] ${s.label}:`, e.message);
      return [];
    }
  }));

  const seen = new Set();
  return per.flat().filter(f => {
    const k = f._driveId + '|' + f.id;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  }).slice(0, 40);
}

// Open a result. Items go to their own page with the detail already
// showing — the same place you'd have got to by clicking through.
async function srchOpenItem(kind, title) {
  const want = String(title || '').toLowerCase();
  const find = list => (list || []).findIndex(x => String(x.Title || x.name || '').toLowerCase() === want);

  if (kind === 'launch' || kind === 'campaign') {
    const isLaunch = kind === 'launch';
    if (typeof showPage === 'function') await showPage(isLaunch ? 'launches' : 'campaigns', isLaunch ? 1 : 2);
    const i = find(isLaunch ? _launchItems : _campaignItems);
    if (i >= 0) (isLaunch ? openLaunchDetail : openCampaignDetail)(i);
    return;
  }

  if (kind === 'event') {
    if (typeof showPage === 'function') await showPage('trade', 3);
    const i = find(_eventFolders);
    if (i >= 0 && typeof openEventFolder === 'function') openEventFolder(i);
    return;
  }

  if (kind === 'training') {
    if (typeof showPage === 'function') await showPage('trade', 3);
    setTimeout(() => {
      const band = document.getElementById('ev-training');
      if (band) band.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 220);
  }
}

// ══════════════════════════════════════════════════════════════
// OPEN EVERYTHING
// ══════════════════════════════════════════════════════════════
//
// 1 Sep 2026, second round. David, for the third time:
//
//   "when you open a campaign or anything like that I want it to open
//    everything in to that sheet like a website … I want it to open
//    everything you click on it and it brings everything open. I've
//    asked for this a couple times now."
//
// The two previous attempts both stopped one click short. First it was
// folder tiles that opened a file list. Then the file list opened a
// reader page. Every version still made you click your way down to a
// document — the structure was on screen, the CONTENT never was.
//
// This is the content. Open a campaign and the whole thing is already
// there, in one scroll: the artwork as pictures, the data sheets and
// PDFs rendered in the page, the videos playing, the decks and Word
// files laid out. Nothing to click to see it. Download and Copy link
// sit on every item for the people who need the file itself, which is
// the half that was always working.
//
// It has to stay fast, so nothing is fetched until it is nearly on
// screen: an IntersectionObserver fills each frame about a screen
// before you reach it. A campaign with sixty files costs the same to
// open as one with three.

const OPEN = { blobs: [], io: null, seq: 0 };

const OE_IMG    = /\.(jpe?g|png|gif|webp|bmp|svg|tiff?)$/i;
const OE_VIDEO  = /\.(mp4|mov|webm|m4v)$/i;
const OE_AUDIO  = /\.(mp3|m4a|wav|ogg)$/i;
const OE_PDF    = /\.pdf$/i;
const OE_OFFICE = /\.(docx?|pptx?|xlsx?|xlsm|potx|dotx)$/i;
const OE_TEXT   = /\.(txt|md|csv|tsv|json|xml|log)$/i;
const OE_LINK   = /\.url$/i;

// 1 Sep 2026 — David: "it seems to take a while to render then some
// things are missing."
//
// The blob route downloads the WHOLE file before a single pixel appears.
// That is right for the full-screen reader, where you asked for one
// document. It is wrong for a page showing ten of them: ten downloads
// start at once, the tab stalls, and whatever loses the race looks
// missing.
//
// So inline, only genuinely small PDFs are fetched as blobs (crisper,
// and quick enough not to notice). Everything else goes through
// Microsoft's preview service, which is one small call and then an
// iframe that paints progressively. The reader still uses the blob for
// everything, unchanged.
const OE_BLOB_MAX  = 1.5 * 1024 * 1024;

function _oeReset() {
  OPEN.blobs.forEach(u => { try { URL.revokeObjectURL(u); } catch (_) {} });
  OPEN.blobs = [];
  if (OPEN.io) { try { OPEN.io.disconnect(); } catch (_) {} OPEN.io = null; }
}

// Walk the whole tree under a folder, keeping the path, so sub-folders
// become sections rather than another thing to click.
async function _oeWalk(driveId, itemId, path, depth, out, cap) {
  if (depth < 0 || out.length >= cap) return;
  let kids;
  try { kids = await fetchDriveChildren(driveId, itemId); }
  catch (e) { console.info('[Open] could not read a folder:', e.message); return; }

  const folders = [];
  for (const k of kids) {
    if (k.folder) folders.push(k);
    else if (out.length < cap) out.push(Object.assign({}, k, { _path: path, _driveId: driveId }));
  }
  // Sequential, not parallel: a campaign folder is a handful of
  // sub-folders and this keeps Graph calls civil.
  for (const f of folders) {
    await _oeWalk(driveId, f.id, path.concat(f.name), depth - 1, out, cap);
  }
}

// Sections read best in the order marketing actually work: the pictures
// first, then the words, then the rest, alphabetically inside each.
const OE_SECTION_RANK = [
  /email\s*camp/i, /image|artwork|photo|visual/i, /social/i, /infographic/i,
  /data\s*sheet|datasheet/i, /blog/i, /website|web\s*page/i, /press|pr\b/i, /video/i,
];
function _oeRank(name) {
  const i = OE_SECTION_RANK.findIndex(rx => rx.test(name));
  return i < 0 ? OE_SECTION_RANK.length : i;
}

function _oeKind(name) {
  if (OE_IMG.test(name))    return 'image';
  if (OE_VIDEO.test(name))  return 'video';
  if (OE_AUDIO.test(name))  return 'audio';
  if (OE_PDF.test(name))    return 'pdf';
  if (OE_OFFICE.test(name)) return 'office';
  if (OE_TEXT.test(name))   return 'text';
  if (OE_LINK.test(name))   return 'link';
  return 'other';
}

function _oeCard(f) {
  const k    = regDoc(f);
  const kind = _oeKind(f.name);
  const ext  = (String(f.name).split('.').pop() || '').toUpperCase().slice(0, 4);
  const meta = [humanSize(f.size), fmtSpDate(f.lastModifiedDateTime)].filter(Boolean).join(' · ');
  return `
    <article class="oe-item oe-${kind}">
      <header class="oe-bar">
        <span class="oe-ext">${escHtml(ext)}</span>
        <span class="oe-name">${escHtml(String(f.name).replace(/\.[a-z0-9]+$/i, ''))}</span>
        <span class="oe-meta">${escHtml(meta)}</span>
        <span class="oe-acts">
          <button class="oe-act" onclick="downloadRegDoc('${k}',event)" title="Download">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          </button>
          <button class="oe-act" onclick="shareRegDoc('${k}',event)" title="Copy link">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>
          </button>
          <button class="oe-act" onclick="openRegDoc('${k}')" title="Open full screen">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="15" height="15"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>
          </button>
        </span>
      </header>
      <div class="oe-stage" data-doc="${k}" data-kind="${kind}">
        <div class="oe-wait"><span class="rdr-spin"></span></div>
      </div>
    </article>`;
}

// The one public entry point. Renders everything under `folderId` into
// `hostId` and returns the number of files it found.
async function openEverything(hostId, driveId, folderId, opts) {
  const host = document.getElementById(hostId);
  if (!host) return 0;
  opts = opts || {};
  _oeReset();

  host.innerHTML = `<div class="oe-boot">
    <div class="skeleton sk-line med"></div>
    <div class="skeleton sk-line"></div>
    <div class="skeleton sk-line short"></div>
  </div>`;

  const files = [];
  await _oeWalk(driveId, folderId, [], (opts.depth === undefined ? 3 : opts.depth),
                files, opts.max || 160);

  if (!files.length) {
    host.innerHTML = `<p class="oe-empty">${escHtml(opts.emptyText
      || 'Nothing is filed in here yet. Anything added in SharePoint shows up here automatically.')}</p>`;
    return 0;
  }

  // Group by the sub-folder each file sits in. Loose files come first,
  // under the item's own name — they used to be invisible entirely.
  const groups = new Map();
  files.forEach(f => {
    const key = (f._path || []).join(' ▸ ');
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(f);
  });

  const keys = [...groups.keys()].sort((a, b) => {
    if (!a) return -1;
    if (!b) return 1;
    return _oeRank(a) - _oeRank(b) || a.localeCompare(b);
  });

  const jump = keys.length > 1
    ? `<nav class="oe-jump">${keys.map((k, i) =>
        `<a href="#oe-s${i}" onclick="_oeJump(event,'oe-s${i}')">${escHtml(k || (opts.rootLabel || 'In this folder'))}<b>${groups.get(k).length}</b></a>`).join('')}</nav>`
    : '';

  host.innerHTML = jump + keys.map((k, i) => {
    const rows = groups.get(k).sort((a, b) => String(a.name).localeCompare(String(b.name)));
    return `
      <section class="oe-sec" id="oe-s${i}">
        <h3 class="oe-sec-head">${escHtml(k || (opts.rootLabel || 'In this folder'))}<span>${rows.length}</span></h3>
        <div class="oe-grid">${rows.map(_oeCard).join('')}</div>
      </section>`;
  }).join('');

  _oeObserve(host);
  console.info(`[Open] ${files.length} file(s) across ${keys.length} section(s) — filling as you scroll.`);
  return files.length;
}

function _oeJump(ev, id) {
  if (ev) ev.preventDefault();
  const el = document.getElementById(id);
  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Fill each frame about a screen before it is reached. Without this a
// campaign with sixty files would fire sixty downloads on open.
function _oeObserve(host) {
  const stages = [...host.querySelectorAll('.oe-stage[data-doc]')];
  if (!stages.length) return;

  if (!('IntersectionObserver' in window)) {
    stages.forEach(_oeFill);          // old browser: just do them all
    return;
  }
  OPEN.io = new IntersectionObserver((entries, io) => {
    entries.forEach(e => {
      if (!e.isIntersecting) return;
      io.unobserve(e.target);
      _oeFill(e.target);
    });
  }, { rootMargin: '600px 0px' });
  stages.forEach(s => OPEN.io.observe(s));
}

function _oeStage(el, html) { if (el) el.innerHTML = html; }

async function _oeFill(el) {
  if (!el || el.dataset.filled) return;
  el.dataset.filled = '1';

  const f = DOCREG[el.dataset.doc];
  if (!f) { _oeStage(el, ''); return; }
  const kind = el.dataset.kind;

  const bail = (msg) => _oeStage(el, `<div class="oe-cant">${escHtml(msg)}
    <button class="oe-alt" onclick="downloadRegDoc('${el.dataset.doc}',event)">Download it ↓</button></div>`);

  try {
    if (kind === 'image') {
      const url = await _rdrDownloadUrl(f);
      if (!url) throw new Error('no download URL');
      _oeStage(el, `<img class="oe-img" src="${escAttr(url)}" alt="${escAttr(f.name)}" loading="lazy"
        onerror="_oeImgFallback(this,'${el.dataset.doc}')">`);

    } else if (kind === 'video') {
      const url = await _rdrDownloadUrl(f);
      if (!url) throw new Error('no download URL');
      _oeStage(el, `<video class="oe-video" src="${escAttr(url)}" controls playsinline preload="metadata"></video>`);

    } else if (kind === 'audio') {
      const url = await _rdrDownloadUrl(f);
      if (!url) throw new Error('no download URL');
      _oeStage(el, `<audio class="oe-audio" src="${escAttr(url)}" controls preload="metadata"></audio>`);

    } else if (kind === 'link') {
      await _oeLink(el, f);

    } else if (kind === 'text') {
      const text = await (await _rdrBlob(f, false)).text();
      _oeStage(el, `<pre class="oe-text">${escHtml(text.slice(0, 40000))}</pre>`);

    } else if (kind === 'pdf') {
      if (Number(f.size) > OE_BLOB_MAX) { await _oePreview(el, f); return; }
      try {
        const url = URL.createObjectURL(await _rdrBlob(f, false));
        OPEN.blobs.push(url);
        _oeStage(el, `<iframe class="oe-frame" src="${escAttr(url)}#view=FitH" title="${escAttr(f.name)}" loading="lazy"></iframe>`);
      } catch (e) {
        console.info('[Open] PDF blob route failed for ' + f.name + ' (' + e.message + ') — using the preview service.');
        await _oePreview(el, f);
      }

    } else if (kind === 'office') {
      // Word, PowerPoint and Excel go straight to the preview service
      // inline. Converting them to PDF through Graph means downloading
      // the whole file first, and a 40 MB launch deck holds up
      // everything under it on the page.
      await _oePreview(el, f);

    } else {
      bail('This file type can’t be shown in the page.');
    }
  } catch (e) {
    console.info('[Open] could not open ' + f.name + ':', e.message);
    await _oePreview(el, f);
  }
}

// A .url shortcut is a two-line ini file — "[InternetShortcut]" and
// "URL=https://…". Three of the Commander Fire Blankets folders (Data
// Sheets, Blog post, Website Pages) hold nothing else, which is why
// they came out as "this file doesn't have a preview we can show you":
// the shortcut was being handed to Microsoft's document previewer,
// which has nothing to preview.
//
// 1 Sep 2026 — and reading it was failing too. The first attempt used
// the pre-authenticated downloadUrl, which is a cross-origin fetch whose
// CORS headers vary by tenant (the same thing that bit the reader in
// August). The Graph /content endpoint with an Authorization header is
// CORS-clean by design, so that goes first now, with the old route as a
// fallback. A shortcut that still can't be read is shown as a link to
// SharePoint rather than a broken preview.
async function _oeLinkTarget(f) {
  try {
    const token = await getAccessToken();
    const res = await fetch(`${GRAPH_BASE}/drives/${f._driveId}/items/${f.id}/content`,
      { headers: { Authorization: 'Bearer ' + token } });
    if (res.ok) {
      const m = (await res.text()).match(/^\s*URL\s*=\s*(\S+)/im);
      if (m) return m[1];
    }
  } catch (e) { console.info('[Open] shortcut via Graph failed for ' + f.name + ':', e.message); }
  try {
    const m = (await (await _rdrBlob(f, false)).text()).match(/^\s*URL\s*=\s*(\S+)/im);
    if (m) return m[1];
  } catch (e) { console.info('[Open] shortcut via downloadUrl failed for ' + f.name + ':', e.message); }
  return '';
}

async function _oeLink(el, f) {
  const raw  = await _oeLinkTarget(f);
  const href = raw ? safeUrl(raw, '') : '';
  const name = String(f.name).replace(/\.url$/i, '');

  if (!href) {
    _oeStage(el, `<div class="oe-cant">This is a shortcut, and its address couldn’t be read.
      <a class="oe-alt" href="${escAttr(safeUrl(f.webUrl, '#'))}" target="_blank" rel="noopener">Open it in SharePoint →</a></div>`);
    return;
  }

  // The file's own name is already in the bar above, so the card shows
  // where the shortcut GOES — which is the thing you actually want to
  // know before clicking it.
  let host = href, path = '';
  try {
    const u = new URL(href);
    host = u.hostname.replace(/^www\./, '');
    path = decodeURIComponent(u.pathname).replace(/\/$/, '');
  } catch (_) {}

  _oeStage(el, `
    <a class="oe-shortcut" href="${escAttr(href)}" target="_blank" rel="noopener" title="${escAttr(name)}">
      <span class="oe-shortcut-ico">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="22" height="22"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
      </span>
      <span class="oe-shortcut-main">
        <span class="oe-shortcut-name">${escHtml(host)}</span>
        <span class="oe-shortcut-host">${escHtml(path || href)}</span>
      </span>
      <span class="oe-shortcut-go">Open →</span>
    </a>`);
}

async function _oePreview(el, f) {
  try {
    const prev = await graphPost(`/drives/${f._driveId}/items/${f.id}/preview`, {});
    const url = prev && prev.getUrl;
    if (!url) throw new Error('no preview URL');
    _oeStage(el, `<iframe class="oe-frame" src="${escAttr(url + (url.includes('?') ? '&' : '?') + 'nb=true')}" title="${escAttr(f.name)}" loading="lazy"></iframe>`);
  } catch (e) {
    console.info('[Open] preview service also failed for ' + f.name + ':', e.message);
    _oeStage(el, `<div class="oe-cant">This one can’t be shown in the page.
      <button class="oe-alt" onclick="downloadRegDoc('${el.dataset.doc}',event)">Download it ↓</button></div>`);
  }
}

async function _oeImgFallback(img, key) {
  const f = DOCREG[key];
  if (!f || !img) return;
  try {
    const url = await driveThumb(f._driveId, f.id, 'large');
    if (url) { img.onerror = null; img.src = url; return; }
  } catch (_) {}
  const stage = img.closest ? img.closest('.oe-stage') : null;
  if (stage) await _oePreview(stage, f);
}

// ══════════════════════════════════════════════════════════════
// NAV MENUS
// ══════════════════════════════════════════════════════════════
//
// 1 Sep 2026 — David: "I'd like that when you hover over the tabs at the
// top for it to show a breakdown so you can click directly in to them."
//
// Each panel is filled the FIRST time you hover it and then kept, and
// everything it needs is already in the five-minute list cache by the
// time anyone gets there — so hovering costs nothing on the second look
// and very little on the first. Clicking a row doesn't just open the
// page: it opens the thing itself.

const NAVM = {};          // key → true once filled

function _navHtml(head, rows, all) {
  if (!rows.length) {
    return `<div class="inh-menu-head">${escHtml(head)}</div>
      <p class="inh-menu-empty">Nothing here yet.</p>
      ${all || ''}`;
  }
  return `<div class="inh-menu-head">${escHtml(head)}</div>${rows.join('')}${all || ''}`;
}

function _navRow(label, onclick, tone, note) {
  return `<button class="inh-menu-item" onclick="${onclick}">
    ${tone ? `<span class="inh-menu-dot ${escAttr(tone)}"></span>` : ''}
    <span class="inh-menu-item-label">${escHtml(label)}</span>
    ${note ? `<span class="inh-menu-n">${escHtml(note)}</span>` : ''}
  </button>`;
}

function _navAll(label, onclick) {
  return `<button class="inh-menu-all" onclick="${onclick}">${escHtml(label)} →</button>`;
}

function _navArg(s) {
  return `'${String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'")}'`;
}

async function navMenuOpen(key) {
  const box = document.getElementById('navm-' + key);
  if (!box || NAVM[key]) return;
  NAVM[key] = true;
  box.innerHTML = '<p class="inh-menu-empty">Loading…</p>';

  try {
    if (key === 'launches' || key === 'campaigns') {
      const isLaunch = key === 'launches';
      const items = await fetchListItems(isLaunch ? HUB_CONFIG.lists.launches : HUB_CONFIG.lists.campaigns);
      const rows = (items || [])
        .filter(f => f.Title)
        .sort((a, b) => String(b.LaunchDate || b.StartDate || '').localeCompare(String(a.LaunchDate || a.StartDate || '')))
        .slice(0, 8)
        .map(f => _navRow(f.Title,
          `navGo('${isLaunch ? 'launch' : 'campaign'}',${_navArg(f.Title)})`,
          ragOf(f.Status),
          f.Status || ''));
      box.innerHTML = _navHtml(isLaunch ? 'Product launches' : 'Campaigns', rows,
        _navAll(isLaunch ? 'See all launches' : 'See all campaigns',
                `navGo('page','${isLaunch ? 'launches' : 'campaigns'}')`));
      return;
    }

    if (key === 'trade') {
      const cats = (HUB_CONFIG.tradeEvents && HUB_CONFIG.tradeEvents.categories) || [];
      const rows = cats.map(c => _navRow(c.label, `navGo('evcat',${_navArg(c.label)})`))
        .concat([_navRow('Training & sessions', "navGo('evcat','__training')")]);
      box.innerHTML = _navHtml('Trade, events & training', rows,
        _navAll('Open the page', "navGo('page','trade')"));
      return;
    }

    if (key === 'training') {
      // The folders as they are in SharePoint — the same ones the page
      // groups by, so what you pick here is what you land on.
      let names = [];
      if (LIB.resources && LIB.resources.loaded) {
        names = [...new Set(LIB.resources.files.map(f => f._cat).filter(Boolean))].sort();
      } else {
        const cfg   = _libCfg('resources');
        const drive = await resolveDrive(HUB_CONFIG.sharepointSite, cfg.library || HUB_CONFIG.documentsLibrary);
        const skip  = (cfg.excludeFolders || []).map(x => String(x).toLowerCase());
        names = (await fetchDriveChildren(drive.id, null))
          .filter(k => k.folder && skip.indexOf(String(k.name).toLowerCase()) < 0)
          .map(k => k.name).sort();
      }
      const rows = names.slice(0, 10).map(nm => _navRow(nm, `navGo('lib',${_navArg(nm)})`));
      box.innerHTML = _navHtml('Resources', rows, _navAll('Open Resources', "navGo('page','training')"));
      return;
    }

    if (key === 'portal') {
      const secs = (HUB_CONFIG.productPortal && HUB_CONFIG.productPortal.sections) || [];
      const rows = secs.map(sec => _navRow(sec.label, `navGo('ppband',${_navArg(sec.key)})`));
      box.innerHTML = _navHtml('Product portal', rows,
        _navAll('Open the Product portal', "navGo('page','portal')"));
      return;
    }
  } catch (e) {
    console.info('[Nav] could not build the "' + key + '" menu:', e.message);
    NAVM[key] = false;      // let the next hover try again
    box.innerHTML = '<p class="inh-menu-empty">Couldn’t load this just now.</p>';
  }
}

// Close the panel the pointer is in, so a click doesn't leave it hanging
// over the page it just opened.
function _navClose() {
  document.querySelectorAll('.inh-nav-item').forEach(el => {
    el.classList.add('nav-shut');
    setTimeout(() => el.classList.remove('nav-shut'), 400);
  });
}

async function navGo(kind, arg) {
  _navClose();

  if (kind === 'page') {
    if (arg === 'portal') { if (typeof openProductPortal === 'function') openProductPortal(); return; }
    const idx = { home: 0, launches: 1, campaigns: 2, trade: 3, training: 4 }[arg];
    if (typeof showPage === 'function') showPage(arg, idx);
    return;
  }

  if (kind === 'launch' || kind === 'campaign') {
    if (typeof srchOpenItem === 'function') await srchOpenItem(kind, arg);
    return;
  }

  if (kind === 'evcat') {
    if (typeof showPage === 'function') await showPage('trade', 3);
    setTimeout(() => {
      if (arg === '__training') {
        const band = document.getElementById('ev-training');
        if (band) band.scrollIntoView({ behavior: 'smooth', block: 'start' });
        return;
      }
      const hit = [...document.querySelectorAll('#ev-index .ev-sec')]
        .find(sec => {
          const t = sec.querySelector('.ev-sec-title');
          return t && t.textContent.trim() === String(arg).trim();
        });
      if (hit) hit.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 400);
    return;
  }

  if (kind === 'lib') {
    if (typeof showPage === 'function') await showPage('training', 4);
    setTimeout(() => {
      const s = LIB.resources;
      if (!s || !s.loaded) return;
      s.tag = 'all'; s.q = ''; s.cat = arg;
      renderLibraryResults('resources');
      const box = document.getElementById('lib-results-resources');
      if (box) box.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 500);
    return;
  }

  if (kind === 'ppband') {
    if (typeof openProductPortal === 'function') await openProductPortal();
    ppOpenSectionByKey(arg);
    return;
  }
}

// Open a Product Portal band by its config key rather than its position,
// because the positions depend on which bands exist today.
function ppOpenSectionByKey(key) {
  const i = (PP_BANDS || []).findIndex(b => b.sec && b.sec.key === key);
  if (i >= 0) { ppOpenSection(i); return; }
  console.info('[Nav] no "' + key + '" band on the portal yet — opening the page instead.');
  const idx = document.getElementById('pp-index');
  if (idx) idx.scrollIntoView({ behavior: 'smooth', block: 'start' });
}
