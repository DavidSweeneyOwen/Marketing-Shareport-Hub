/**
 * CheckFire Marketing Hub — App Logic
 * ─────────────────────────────────────
 * Navigation, UI interactions, and live data rendering.
 */

// ─── Page Navigation ──────────────────────────────────────────

const PAGE_KEYS = ['home', 'launches', 'campaigns', 'trade', 'training', 'portal'];
const dataLoaded = {};

// 18 SEP 2026 — David: "when you open a product launch or campaign it
// always takes you back to the one you were in … when you click out and
// click back in it takes you to the home page in that tab."
//
// Every page in the hub has a front and a drill-down, and the drill-down
// was left open when you navigated away: the detail panel is a
// display:none swap inside the page, and nothing ever put it back. So
// the tab remembered a launch from twenty minutes ago and opened on it.
//
// Clicking a tab now lands on that tab's front, every time. Each page
// already had a close function — this is the one place that calls them,
// so a new page means one more line here and nothing else.
//
// GUARDED ON LOADED. Resetting a page whose data has never arrived
// would hide the skeleton the loader is about to fill, so a page that
// hasn't loaded is left exactly as it is.
function resetPageView(id) {
  try {
    // ONE EXCEPTION. closeReader() calls showPage() to put the reader
    // back where it came from, and where it came from is quite often an
    // open campaign or an open portal section. Coming out of the reader
    // or the search page is a RETURN, not a tab click, so it leaves the
    // page exactly as it found it.
    const from = document.querySelector('.page.active');
    const fromId = from ? String(from.id).replace(/^page-/, '') : '';
    if (fromId === 'reader' || fromId === 'search') return;

    if (id === 'launches'  && typeof closeLaunchDetail   === 'function') closeLaunchDetail();
    if (id === 'campaigns' && typeof closeCampaignDetail === 'function') closeCampaignDetail();
    if (id === 'trade'     && typeof closeEventFolder    === 'function') closeEventFolder();

    const lib = (typeof LIB === 'undefined') ? null : LIB;
    if (id === 'training' && typeof libFoldersBack === 'function'
        && lib && lib.resources && lib.resources.loaded) libFoldersBack('resources');
    if (id === 'portal' && typeof ppCloseSection === 'function'
        && lib && lib.product && lib.product.loaded) ppCloseSection();
  } catch (e) {
    console.info('[Nav] could not reset the ' + id + ' page: ' + e.message);
  }
}

async function showPage(id, idx) {
  resetPageView(id);
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  const el = document.getElementById('page-' + id);
  if (el) el.classList.add('active');

  document.querySelectorAll('.nav-link').forEach((a, i) => {
    a.classList.toggle('active', i === idx);
  });
  if (typeof updateNavActive === 'function') updateNavActive(id);
  window.scrollTo({ top: 0, behavior: 'smooth' });

  if (id === 'home')     setTimeout(animateBars, 250);
  if (id === 'training') setTimeout(animateTrainingRing, 250);

  // Load data for page if not yet loaded (and not in demo mode)
  if (!window.HUB_DEMO_MODE && !dataLoaded[id]) {
    dataLoaded[id] = true;
    await loadPageData(id);
  }
}

async function loadPageData(pageId) {
  switch (pageId) {
    case 'home':      await loadHomeData();      break;
    // Launches / Campaigns / Events all come from loadSharePointData()
    // in graph.js (fetchListItems + renderers, 5-min cache). The old
    // per-page loaders below (loadLaunchData etc.) call functions that
    // no longer exist and must NOT be wired back in.
    case 'launches':
    case 'campaigns': await loadSharePointData(); break;
    // Trade & Events reads the Documents ▸ Events folders, not a list.
    case 'trade':     await loadTradeEvents();    break;
    case 'training':  await loadResourcesData(); break;
    // Product Portal is its own page now — Resources is just the
    // Marketing Library, per marketing's Aug feedback.
    case 'portal':    await loadProductPortal();  break;
    // 'reader' and 'search' are filled by whatever navigated to them —
    // openDocFile() and openSiteSearch() — so they have no loader here.
  }
}

// ─── Home Page Data ───────────────────────────────────────────

async function loadHomeData() {
  renderQuickLinks();
  await Promise.all([
    loadNotices(),
    loadHeroNews(),
    loadBlogsCarousel(),
    loadLandingPages(),
    loadHeroLaunch(),
    startCountdown(),
    loadHomeVideos(),
    loadWall(),
    loadTraining(),
    loadPolls(),
    loadUpdatesDock(),
  ]);
}

// ─── Quick Links ──────────────────────────────────────────────
// Renders HUB_CONFIG.quickLinks. Needs no network, so it runs first
// and the box is never empty while something else is loading. Marketing
// edit the list in config.js; nothing here needs changing to add one.
function renderQuickLinks() {
  const box = document.getElementById('home-quicklinks');
  if (!box) return;
  const links = (HUB_CONFIG && HUB_CONFIG.quickLinks) || [];
  box.innerHTML = links.map(l => `
    <a class="inh-ql" href="${escAttr(safeUrl(l.url))}" target="_blank" rel="noopener">
      <span class="inh-ql-icon" style="background:${escAttr(l.colour || '#111')}">${escHtml(l.initials || '')}</span>${escHtml(l.label || '')}
    </a>`).join('');
}

// ─── Horizontal carousels (blogs / landing pages) ─────────────
// Scrolls the track by ~one card width. Buttons wired in index.html.
function scrollCarousel(trackId, dir) {
  const track = document.getElementById(trackId);
  if (!track) return;
  const card = track.querySelector('.cara-card, .cara-skel, .train-card');
  const step = card ? card.getBoundingClientRect().width + 16 : 316;
  track.scrollBy({ left: dir * step * 1.5, behavior: 'smooth' });
}

// ─── Active nav highlight ─────────────────────────────────────
// ui.js's showPage() calls updateNavActive(id) if it exists. Map the
// page id to its nav link so the underline follows the current page.
function updateNavActive(id) {
  const map = { home:'navl-home', launches:'navl-launches', campaigns:'navl-campaigns', trade:'navl-trade', training:'navl-training', portal:'navl-portal' };
  document.querySelectorAll('.inh-nav-link').forEach(a => a.classList.remove('active'));
  const el = document.getElementById(map[id]);
  if (el) el.classList.add('active');
}

// ─── Product Portal ───────────────────────────────────────────
// Aug 2026: Resources was cut back to the Marketing Library alone, so
// the Product Portal is a page in its own right rather than a tab.
// Files still open in-hub — the browser instance and its element ids
// are unchanged, only where they live on the page.
async function openProductPortal() {
  await showPage('portal', 5);
  if (typeof loadProductPortal === 'function') loadProductPortal();
  if (typeof updateNavActive === 'function') updateNavActive('portal');
}

// ─── WordPress News ───────────────────────────────────────────

// Build one horizontal carousel card from a WordPress post/page object.
function _caraCard(item, idx) {
  const link  = safeUrl(item.link);
  const image = safeCssUrl(item.image);
  // No featured image (most WordPress *pages* have none) — drop the image
  // area entirely rather than leaving an empty grey box, and let the CSS
  // give the card a red rule instead. loadLandingPages() may then slot a
  // picture in from SharePoint, which is why the card carries an id.
  return `
    <a class="cara-card${image ? '' : ' no-img'}" ${idx === undefined ? '' : `id="cara-${idx}"`} href="${escAttr(link)}" target="_blank" rel="noopener">
      ${image ? `<div class="cara-img" style="background-image:url('${image}')"></div>` : ''}
      <div class="cara-body">
        <div class="cara-date">${escHtml(item.date)}</div>
        <div class="cara-title">${escHtml(item.title)}</div>
        ${item.excerpt ? `<p class="cara-excerpt">${escHtml(item.excerpt)}</p>` : ''}
      </div>
    </a>`;
}

// Latest Blogs carousel — WordPress posts.
async function loadBlogsCarousel() {
  const track = document.getElementById('home-blogs-track');
  if (!track) return;
  try {
    const posts = await fetchWordPressNews();
    if (!posts.length) { track.innerHTML = '<p class="prose dim">No blog posts found.</p>'; return; }
    // No index passed — only the landing-page cards need ids (the
    // SharePoint artwork is slotted into them by id later).
    track.innerHTML = posts.map(p => _caraCard(p)).join('');
  } catch (e) {
    track.innerHTML = `<p class="sp-error">Couldn't load blogs: ${escHtml(e.message)}</p>`;
  }
}

// Updated Landing Pages carousel — WordPress pages, newest-modified.
// Section hides itself if the pages endpoint returns nothing.
//
// 26 Aug 2026: marketing are putting artwork for these pages into
// Documents ▸ Images for Landing Pages, matched to a page by filename
// (see config.js). The carousel renders straight away with the
// text-only cards and takes the pictures when SharePoint answers, so a
// slow library never holds the home page up. Any page with no matching
// image keeps the text-only card — never an empty grey box.
// 10 Sep 2026 — marketing: "Can those four landing pages be first in
// that section, as I know that's the general one? The general one can
// go after those when people scroll to the left."
//
// WordPress orders these by last-modified, so editing an evergreen page
// pushes the product ones off the front. HUB_CONFIG.wordpress.pinned
// names the pages that lead, in the order marketing want them; anything
// not named keeps the newest-first order behind them. Marketing can
// re-order the carousel by editing that list — no code change.
function _orderLandingPages(pages) {
  const pinned = (HUB_CONFIG.wordpress && HUB_CONFIG.wordpress.pinned) || [];
  if (!pinned.length || !pages || !pages.length) return pages || [];
  const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  const rank = p => {
    const t = norm(p.title), l = norm(p.link);
    for (let i = 0; i < pinned.length; i++) {
      const k = norm(pinned[i]);
      if (k && (t === k || t.includes(k) || k.includes(t) || l.includes(k))) return i;
    }
    return pinned.length;          // everything else keeps its order, after
  };
  return pages
    .map((p, i) => ({ p, r: rank(p), i }))
    .sort((a, b) => (a.r - b.r) || (a.i - b.i))
    .map(x => x.p);
}

async function loadLandingPages() {
  const section = document.getElementById('home-pages-section');
  const track   = document.getElementById('home-pages-track');
  if (!track) return;
  try {
    const pages = _orderLandingPages(await fetchWordPressPages());
    if (!pages.length) { if (section) section.style.display = 'none'; return; }
    track.innerHTML = pages.map((p, i) => _caraCard(p, i)).join('');
    if (section) section.style.display = '';

    if (typeof fetchLandingImages === 'function') {
      const images = await fetchLandingImages();
      if (!images.length) return;
      // 10 Sep 2026 — assign across ALL pages at once so one picture can
      // never be handed to two cards. See _landingScore in graph.js.
      const picked = typeof assignLandingImages === 'function'
        ? assignLandingImages(images, pages)
        : new Map(pages.map((p, i) => [i, matchLandingImage(images, p)]));
      pages.forEach((p, i) => {
        const url = picked.get(i);
        if (!url) return;
        const card = document.getElementById('cara-' + i);
        if (!card) return;
        card.classList.remove('no-img');
        // 2 Sep 2026 — marketing, on the Flat-Pack Tubular Stand card:
        // "something happened here". Two pictures were stacked inside
        // one card, overflowing it. This used to insert a NEW .cara-img
        // every time without checking, so any page that had BOTH a
        // WordPress featured image and a SharePoint match ended up with
        // two. Reuse the one that's already there.
        let media = card.querySelector('.cara-img');
        if (!media) {
          media = document.createElement('div');
          media.className = 'cara-img';
          card.insertBefore(media, card.firstElementChild);
        }
        media.style.backgroundImage = `url('${safeCssUrl(url)}')`;
      });
    }
  } catch (e) {
    console.warn('Landing pages unavailable:', e.message);
    if (section) section.style.display = 'none';
  }
}

async function loadHeroNews() {
  try {
    const posts = await fetchWordPressNews();
    if (!posts.length) return;
    const feature = posts[0];
    const featureEl = document.getElementById('hero-feature');
    if (featureEl) {
      featureEl.href = safeUrl(feature.link);
      const img = document.getElementById('hero-feature-img');
      const featureImg = safeCssUrl(feature.image);
      if (img && featureImg) img.style.backgroundImage = `url('${featureImg}')`;
      const title = document.getElementById('hero-feature-title');
      if (title) title.textContent = feature.title;
      const date = document.getElementById('hero-feature-date');
      if (date) date.textContent = feature.date;
      featureEl.classList.remove('skeleton-card');
    }
    const sideContainer = document.getElementById('hero-side-articles');
    if (sideContainer && posts.length > 1) {
      sideContainer.innerHTML = posts.slice(1, 4).map(post => {
        const link  = safeUrl(post.link);
        const image = safeCssUrl(post.image);
        return `
        <a class="inh-hero-side-item" href="${escAttr(link)}" target="_blank" rel="noopener">
          <div class="inh-hero-img" ${image ? `style="background-image:url('${image}')"` : 'style="background:#2A2A2A"'}></div>
          <div class="inh-hero-overlay">
            <span class="inh-tag blue" style="font-size:9px">News</span>
            <div class="inh-hero-title">${escHtml(post.title)}</div>
            <div class="inh-hero-date">${escHtml(post.date)}</div>
          </div>
        </a>`;
      }).join('');
    }
  } catch (e) {
    console.warn('Hero news load failed:', e.message);
  }
}

// Hero "Upcoming Launch" box — shows the next upcoming launch (or the
// most recent if none are in the future), from the Product Launches list.
function loadHeroLaunch() {
  const el = document.getElementById('home-hero-launch-body');
  if (!el) return;
  return fetchListItems(HUB_CONFIG.lists.launches).then(items => {
    if (!items.length) {
      el.innerHTML = '<p class="prose dim">No product launches in SharePoint yet.</p>';
      return;
    }
    const now = new Date();
    const withDate = items.filter(f => f.LaunchDate && !isNaN(new Date(f.LaunchDate)));
    const upcoming = withDate
      .filter(f => new Date(f.LaunchDate) >= now)
      .sort((a, b) => String(a.LaunchDate).localeCompare(String(b.LaunchDate)));
    const past = withDate
      .sort((a, b) => String(b.LaunchDate).localeCompare(String(a.LaunchDate)));
    const f = upcoming[0] || past[0] || items[0];
    const isFuture = f.LaunchDate && new Date(f.LaunchDate) >= now;
    el.innerHTML = `
      <div class="hbox-launch" onclick="showPage('launches',1)">
        <span class="hbox-launch-tag">${escHtml(f.Status || (isFuture ? 'Upcoming' : 'Latest'))}</span>
        <div class="hbox-launch-title">${escHtml(f.Title || 'Untitled')}</div>
        <div class="hbox-launch-date">${f.LaunchDate ? formatDate(f.LaunchDate) : ''}</div>
      </div>
      <a class="hbox-more" onclick="showPage('launches',1)">View all launches →</a>`;
  }).catch(e => {
    console.warn('Hero launch unavailable:', e.message);
    el.innerHTML = '<p class="prose dim">Sign in to see the latest product launch.</p>';
  });
}

// Next Major Event countdown — driven by the SharePoint Events list.
// Card stays hidden unless there is an upcoming event with a date.
let _countdownTimer = null;

function startCountdown() {
  const card = document.getElementById('home-countdown');
  if (!card) return;
  return fetchListItems(HUB_CONFIG.lists.events).then(items => {
    const now = new Date();
    const upcoming = items
      .filter(f => f.EventDate && !isNaN(new Date(f.EventDate)) && new Date(f.EventDate) >= now)
      .sort((a, b) => String(a.EventDate).localeCompare(String(b.EventDate)))[0];
    if (!upcoming) { card.style.display = 'none'; return; }

    const nameEl = document.getElementById('countdown-name');
    const dateEl = document.getElementById('countdown-date');
    if (nameEl) nameEl.textContent = upcoming.Title || 'Untitled';
    if (dateEl) dateEl.textContent = [formatDate(upcoming.EventDate), upcoming.Location].filter(Boolean).join(' · ');
    card.style.display = '';

    const target = new Date(upcoming.EventDate);
    function tick() {
      const diff = target - new Date();
      if (diff <= 0) return;
      const d = document.getElementById('cd-days');
      const h = document.getElementById('cd-hours');
      const m = document.getElementById('cd-mins');
      if (d) d.textContent = String(Math.floor(diff / 86400000)).padStart(2, '0');
      if (h) h.textContent = String(Math.floor((diff % 86400000) / 3600000)).padStart(2, '0');
      if (m) m.textContent = String(Math.floor((diff % 3600000) / 60000)).padStart(2, '0');
    }
    tick();
    clearInterval(_countdownTimer);
    _countdownTimer = setInterval(tick, 30000);
  }).catch(e => {
    console.warn('Countdown unavailable:', e.message);
    card.style.display = 'none';
  });
}

// (Demo-era loaders removed 7 Jul 2026 — every page now renders live
// SharePoint data via graph.js. Do not re-add hardcoded content here.)

// ─── Animation helpers (preserved from prototype) ────────────

function animateBars() {
  document.querySelectorAll('.bar span').forEach(b => {
    const pct = b.getAttribute('data-pct');
    b.style.width = '0%';
    requestAnimationFrame(() => {
      setTimeout(() => { b.style.width = pct + '%'; }, 80);
    });
  });
}

function animateTrainingRing() {
  const ring = document.getElementById('trainingRing');
  if (!ring) return;
  const circumference = 2 * Math.PI * 34;
  const pct = 12 / 14;
  ring.style.strokeDashoffset = circumference;
  requestAnimationFrame(() => {
    setTimeout(() => {
      ring.style.transition = 'stroke-dashoffset 1.4s cubic-bezier(0.22,1,0.36,1)';
      ring.style.strokeDashoffset = circumference * (1 - pct);
    }, 100);
  });
}

// ─── UI helpers ───────────────────────────────────────────────

function setSkeleton(container, count, type = 'sk-line') {
  container.innerHTML = Array.from({length: count}, () =>
    `<div class="skeleton ${type}"></div>`
  ).join('');
}

function setMetric(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function showToast(msg) {
  let t = document.getElementById('hub-toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'hub-toast';
    t.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#0A0A0A;color:#fff;padding:10px 20px;border-radius:10px;font-size:13px;font-weight:500;z-index:9998;opacity:0;transition:opacity .2s;white-space:nowrap;pointer-events:none';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.style.opacity = '1';
  clearTimeout(t._tid);
  t._tid = setTimeout(() => { t.style.opacity = '0'; }, 2400);
}

function filterWhats(cat, btn) {
  document.querySelectorAll('.wf-pill').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.whats-item').forEach(item => {
    item.style.display = (cat === 'all' || item.dataset.cat === cat) ? '' : 'none';
  });
}

function switchTrainingTab(btn, tab) {
  document.querySelectorAll('.training-tab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.ltab-pane').forEach(p => p.classList.remove('active'));
  const pane = document.getElementById('ttab-' + tab);
  if (pane) pane.classList.add('active');
}

// (The old demo answerPoll() stub was removed — real voting is
// votePoll() in graph.js, which writes to the SharePoint Poll Votes list.)

function postToWall(btn) {
  const ta = btn.closest('.wall-composer')?.querySelector('.wall-comp-text');
  if (!ta?.value?.trim()) { showToast('Write something first'); return; }
  showToast('Posted to the wall');
  ta.value = '';
}

// ─── Utility ─────────────────────────────────────────────────

function escHtml(str) {
  return String(str || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function ucFirst(str) {
  return str ? str.charAt(0).toUpperCase() + str.slice(1) : '';
}

function formatDate(dateStr) {
  if (!dateStr) return '—';
  return new Date(dateStr).toLocaleDateString('en-GB', { day:'numeric', month:'short', year:'numeric' });
}

function colourForStatus(status) {
  const s = (status || '').toLowerCase();
  if (s === 'live') return 'red';
  if (s === 'planning') return 'amber';
  return 'grey';
}

function dotColour(status) {
  const s = (status || '').toLowerCase();
  if (s === 'live') return 'green';
  if (s === 'planning') return 'amber';
  return 'grey';
}

function folderIcon() {
  return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="18" height="18"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>`;
}

// ─── Header search ───────────────────────────────────────────
// The search box in the header has been decorative since the first
// build — it accepted typing and did nothing, which is worse than not
// being there ("you also can't search anything", 26 Aug). It now hands
// the query to Ember, who searches all three SharePoint sites and
// answers in the corner panel. Enter or the icon runs it.
// 26 Aug 2026 (fix 4): this used to hand the query to Ember. David:
// "I'd rather the search bar be able to search the site not talk to
// ember through it." It now runs a real search over the hub's own
// content — launches, campaigns, events, training and both document
// libraries — and lands on the results page. Ember keeps her own
// button for the questions that need a document read properly.
//
// The query stays in the box, because after a search people narrow it
// rather than start again.
function initHeaderSearch() {
  const input = document.getElementById('hub-search-input');
  if (!input) return;
  input.placeholder = 'Search the hub — press Enter…';
  input.addEventListener('keydown', e => {
    if (e.key !== 'Enter') return;
    const q = input.value.trim();
    if (!q) return;
    input.blur();
    if (typeof openSiteSearch === 'function') openSiteSearch(q);
  });
}

// ─── Sticky-header height ────────────────────────────────────
// The filter rail on Launches and Campaigns sticks BELOW the hub's
// own sticky header. That header's height depends on the viewport, so
// publish it as a CSS variable rather than guessing a number.
function syncHeaderHeight() {
  const h = document.getElementById('intranet-header');
  if (!h) return;
  document.documentElement.style.setProperty('--hub-header-h', h.offsetHeight + 'px');
}
window.addEventListener('resize', syncHeaderHeight);
window.addEventListener('load', syncHeaderHeight);

// ─── Init ────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  syncHeaderHeight();
  initHeaderSearch();
  const launchDays = document.getElementById('launchDays');
  if (launchDays) launchDays.textContent = '—';

  const authed = await initAuth();
  if (authed || window.HUB_DEMO_MODE) {
    await loadPageData('home');
    animateBars();
    setTimeout(animateBars, 400);
    // Now we're signed in, pull showroom bookings from SharePoint
    // (the calendar's "who's coming in"). Safe to call again — it
    // just re-renders with the user's token.
    if (typeof loadShowroomData === 'function') loadShowroomData();
  }
});

/* ─── Browser history — 17 Sep 2026 ────────────────────────────
 * Until today the hub was one HTML page that swapped a .page class,
 * so nothing a user clicked ever reached the browser's own history:
 * Back took them OFF the site, and marketing.checkfire.co.uk was the
 * only URL there was.
 *
 * This wraps the functions that move the user, records each one as a
 * history entry, and plays it back on popstate. Nothing calls into
 * here — the wrapping is done at DOMContentLoaded, so every existing
 * onclick in index.html and every link graph.js builds gets history
 * for free. Adding a new navigation function means adding its NAME to
 * one of the two lists below; there is nothing else to wire up.
 *
 * Top-level pages also get a real path (/campaigns, /product-portal),
 * which works because staticwebapp.config.json falls back to
 * index.html for any path that isn't a file. Drill-downs keep their
 * page's path and carry their state in the history entry.
 * ───────────────────────────────────────────────────────────── */
(function () {
  'use strict';

  // Page ids are the legacy ones in index.html — 'training' is the
  // Resources page and 'portal' is the Product Portal.
  var PAGE_PATH = {
    home: '/', launches: '/launches', campaigns: '/campaigns',
    trade: '/trade-events', training: '/resources', portal: '/product-portal',
    reader: '/document', search: '/search'
  };
  var PATH_PAGE = {};
  Object.keys(PAGE_PATH).forEach(function (id) { PATH_PAGE[PAGE_PATH[id]] = id; });

  // Page-level: these land the user on a whole page.
  var PAGE_FNS = ['showPage', 'openProductPortal'];
  // Deep: these open something inside the page the user is already on.
  var DEEP_FNS = [
    'openCampaignDetail', 'closeCampaignDetail', 'openLaunchDetail', 'closeLaunchDetail',
    'ppOpenSection', 'ppCloseSection', 'libCat', 'libPick', 'libReset', 'fbCrumb',
    'openEventFolder', 'openDetailSubfolder', 'openDetailAsset',
    'srchOpenItem', 'updOpenItem', 'openSiteSearch', 'openDocFile',
    // 18 Sep 2026 — the Resources folder cards.
    'libOpenFolder', 'libFoldersBack'
  ];

  var suppress = false;  // true while we are re-rendering FROM history
  var pending  = null;
  var timer    = null;

  function currentPage() {
    var el = document.querySelector('.page.active');
    return el ? el.id.replace(/^page-/, '') : 'home';
  }

  // Only primitives survive into history.state. A DOM node becomes null
  // — every call site passes it as "the button to highlight", which the
  // renderer works out again anyway. Anything else (openDocFile's file
  // object) marks the view unrestorable: we still record it so Back has
  // somewhere to come from, we just land on its page rather than reopen
  // it. Returns null when the args can't be trusted.
  function safeArgs(args) {
    var ok = true;
    var out = [].map.call(args, function (a) {
      if (a === null || a === undefined) return null;
      var t = typeof a;
      if (t === 'string' || t === 'number' || t === 'boolean') return a;
      if (typeof Node !== 'undefined' && a instanceof Node) return null;
      ok = false;
      return null;
    });
    return ok ? out : null;
  }

  function key(d) {
    return d.page + (d.deep ? '/' + d.fn + '(' + JSON.stringify(d.args) + ')' : '');
  }

  function record(fn, args, deep) {
    if (suppress) return;
    var safe = safeArgs(args);
    pending = {
      fn: fn, args: safe || [], deep: deep,
      page: currentPage(), restorable: safe !== null
    };
    if (!timer) timer = setTimeout(flush, 0);
  }

  function flush() {
    timer = null;
    var d = pending;
    pending = null;
    if (!d) return;
    d.page = currentPage();                       // it has settled by now
    var url = PAGE_PATH[d.page] || location.pathname;
    var cur = history.state && history.state.hub;
    // One click is one entry. The Product Portal nav link calls
    // showPage('portal') and then openProductPortal(), which is two
    // records for the same view — the second replaces the first.
    if (cur && key(cur) === key(d)) history.replaceState({ hub: d }, '', url);
    else history.pushState({ hub: d }, '', url);
  }

  function wrap(name, deep) {
    var orig = window[name];
    if (typeof orig !== 'function') return;
    window[name] = function () {
      var r = orig.apply(this, arguments);
      record(name, arguments, deep);
      return r;
    };
  }

  function replay(d) {
    suppress = true;
    try {
      // 18 Sep 2026 — a deep entry records the view INSIDE a page, and
      // used to be replayed straight onto whatever page the reader
      // happened to be on: Back out of a campaign detail into Trade &
      // Events and then Back again re-opened the detail with the events
      // page still on screen. Put its page back first. showPage resets
      // that page to its front, and the call below then re-opens the
      // one thing the entry is actually for.
      if (d && d.deep && d.page && typeof showPage === 'function') showPage(d.page);
      if (d && d.restorable && typeof window[d.fn] === 'function') {
        window[d.fn].apply(null, d.args);
      } else if (typeof showPage === 'function') {
        showPage((d && d.page) || 'home');
      }
    } finally {
      setTimeout(function () { suppress = false; }, 0);
    }
  }

  // A deep link has to wait for the token, or its page renders empty.
  function whenReady(cb) {
    var tries = 0;
    (function tick() {
      if (window.HUB_DEMO_MODE || (window.AUTH && window.AUTH.token)) return cb();
      if (++tries > 60) return;                   // ~30s, then leave it alone
      setTimeout(tick, 500);
    })();
  }

  function install() {
    PAGE_FNS.forEach(function (n) { wrap(n, false); });
    DEEP_FNS.forEach(function (n) { wrap(n, true); });

    var path  = location.pathname.replace(/\/+$/, '') || '/';
    var start = PATH_PAGE[path] || 'home';

    history.replaceState(
      { hub: { fn: 'showPage', args: [start], deep: false, page: start, restorable: true } },
      '', PAGE_PATH[start] || '/'
    );

    // /campaigns typed straight into the address bar opens Campaigns.
    if (start !== 'home' && start !== 'reader' && start !== 'search') {
      whenReady(function () {
        suppress = true;
        try {
          if (start === 'portal' && typeof openProductPortal === 'function') openProductPortal();
          else if (typeof showPage === 'function') showPage(start);
        } finally {
          setTimeout(function () { suppress = false; }, 0);
        }
      });
    }

    window.addEventListener('popstate', function (e) {
      replay(e.state && e.state.hub);
    });
  }

  // Installed on DOMContentLoaded, or straight away if that has already
  // fired — app.js must keep working whether or not it is deferred.
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', install);
  } else {
    install();
  }
})();
