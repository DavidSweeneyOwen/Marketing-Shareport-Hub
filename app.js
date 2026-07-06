/**
 * CheckFire Marketing Hub — App Logic
 * ─────────────────────────────────────
 * Navigation, UI interactions, and live data rendering.
 */

// ─── Page Navigation ──────────────────────────────────────────

const PAGE_KEYS = ['home', 'launches', 'campaigns', 'trade', 'training'];
const dataLoaded = {};

async function showPage(id, idx) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  const el = document.getElementById('page-' + id);
  if (el) el.classList.add('active');

  document.querySelectorAll('.nav-link').forEach((a, i) => {
    a.classList.toggle('active', i === idx);
  });
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
    case 'campaigns':
    case 'trade':     await loadSharePointData(); break;
    case 'training':  await loadResourcesData(); break;
  }
}

// ─── Home Page Data ───────────────────────────────────────────

async function loadHomeData() {
  await Promise.all([
    loadHeroNews(),
    loadProductNewsList(),
    startCountdown(),
  ]);
  try {
    const files = await fetchRecentFiles();
    renderRecentFiles(files);
  } catch (e) {
    console.warn('Recent files unavailable:', e.message);
  }
}

// ─── WordPress News ───────────────────────────────────────────

async function loadWordPressNews() {
  const container = document.getElementById('sp-wp-news');
  if (!container) return;
  container.innerHTML = `<div class="wp-news-grid">${Array(3).fill('<div class="wp-news-card skeleton" style="height:200px;border-radius:12px"></div>').join('')}</div>`;
  try {
    const posts = await fetchWordPressNews();
    if (!posts.length) { container.innerHTML = '<p class="sp-error">No news posts found.</p>'; return; }
    container.innerHTML = `<div class="wp-news-grid">${posts.map(post => {
      const link  = safeUrl(post.link);
      const image = safeCssUrl(post.image);
      return `
      <a class="wp-news-card" href="${escAttr(link)}" target="_blank" rel="noopener">
        ${image ? `<div class="wp-news-img" style="background-image:url('${image}')"></div>`
                : `<div class="wp-news-img wp-news-img-placeholder"></div>`}
        <div class="wp-news-body">
          <div class="wp-news-date">${escHtml(post.date)}</div>
          <h3 class="wp-news-title">${escHtml(post.title)}</h3>
          <p class="wp-news-excerpt">${escHtml(post.excerpt)}</p>
        </div>
      </a>`;
    }).join('')}</div>`;
  } catch (e) {
    container.innerHTML = `<p class="sp-error">Couldn't load news: ${escHtml(e.message)}</p>`;
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

function loadProductNewsList() {
  const el = document.getElementById('home-product-news');
  if (!el) return;
  const items = [
    { tag: 'New Launch',    name: 'ProGuard 6kg CO₂ Extinguisher', date: 'Launching 2 June 2026',  onclick: "showPage('launches',1)" },
    { tag: 'Coming Soon',   name: 'FireShield Premium Kit',         date: 'Launching 19 June 2026', onclick: "showPage('launches',1)" },
    { tag: 'Back in Stock', name: 'Commander Stand',                date: 'Available now',           onclick: "showPage('launches',1)" },
  ];
  el.innerHTML = items.map(item => `
    <div class="inh-product-item" onclick="${item.onclick}">
      <div class="inh-product-img-placeholder">CF</div>
      <div class="inh-product-info">
        <div class="inh-product-tag">${escHtml(item.tag)}</div>
        <div class="inh-product-name">${escHtml(item.name)}</div>
        <div class="inh-product-date">${escHtml(item.date)}</div>
      </div>
    </div>
  `).join('');
}

function startCountdown() {
  const target = new Date('2026-07-08T09:00:00');
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
  setInterval(tick, 30000);
}

function renderRecentFiles(files) {
  const container = document.getElementById('sp-recent-files');
  if (!container || !files.length) return;
  container.innerHTML = files.map(f => `
    <a class="sp-doc-card" href="${f.url}" target="_blank" rel="noopener">
      <div class="sp-doc-icon">${fileIcon()}</div>
      <div>
        <div class="sp-doc-name">${escHtml(f.name)}</div>
        <div class="sp-doc-meta">${f.modified}</div>
      </div>
    </a>
  `).join('');
}

// ─── Launches Page Data ───────────────────────────────────────

async function loadLaunchData() {
  const container = document.getElementById('sp-launches-list');
  if (!container) return;
  setSkeleton(container, 3);

  // Try structured list first; fall back to document folder
  let launches = [];
  let usedFallback = false;
  try {
    launches = await fetchLaunches();
  } catch (e) {
    console.warn('Launches list unavailable, trying document folder:', e.message);
  }

  if (launches.length) {
    // Rich list-item cards from SharePoint List
    const featured = launches.find(l => l.status.toLowerCase() !== 'completed') || launches[0];
    updateLaunchHero(featured);

    container.innerHTML = launches.map(l => {
      const days = daysUntil(l.launchDate);
      const badge = statusBadgeClass(l.status);
      return `
        <div class="sp-list-item">
          <div>
            <div class="sp-list-title">${escHtml(l.title)}</div>
            <div class="sp-list-sub">${l.sku ? 'SKU: ' + l.sku + ' · ' : ''}${l.launchDate ? formatDate(l.launchDate) : '—'}</div>
          </div>
          <span class="sp-badge ${badge}">${days !== null && badge !== 'completed' ? days + 'd' : ucFirst(l.status)}</span>
        </div>
      `;
    }).join('');
  } else {
    // Fallback: show files/folders from the Launches document folder
    try {
      const folder = HUB_CONFIG.folders?.launches || 'Launches';
      const docs = await fetchDocuments(folder);
      if (!docs.length) { container.innerHTML = '<p class="sp-error">No launches found in SharePoint.</p>'; return; }
      container.innerHTML = docs.map(d => `
        <a class="sp-list-item" href="${escAttr(d.url)}" target="_blank" rel="noopener" style="text-decoration:none;color:inherit">
          <div>
            <div class="sp-list-title">${d.isFolder ? folderIcon() + ' ' : fileIcon(d.ext) + ' '}${escHtml(d.name)}</div>
            <div class="sp-list-sub">Updated ${escHtml(d.modified)}</div>
          </div>
          <span class="sp-badge upcoming">View</span>
        </a>
      `).join('');
    } catch (e2) {
      container.innerHTML = `<p class="sp-error">Couldn't load launches: ${escHtml(e2.message)}</p>`;
    }
  }
}

function updateLaunchHero(launch) {
  const titleEl = document.querySelector('.lhero-title');
  const ledeEl  = document.querySelector('.lhero-lede');
  const badgeEl = document.querySelector('.lhero-lead .badge');
  const daysEl  = document.getElementById('launchDays');

  if (titleEl) titleEl.textContent = launch.title;
  if (ledeEl && launch.description) ledeEl.textContent = launch.description;
  if (daysEl && launch.launchDate) daysEl.textContent = daysUntil(launch.launchDate);
  if (badgeEl && launch.launchDate) {
    const d = daysUntil(launch.launchDate);
    badgeEl.innerHTML = `<span class="status-dot ${d > 0 ? 'amber' : 'green'}"></span> ${d > 0 ? 'Launching ' + formatDate(launch.launchDate) + ' · ' + d + ' days to go' : 'Live now'}`;
  }
}

// ─── Campaigns Page Data ──────────────────────────────────────

async function loadCampaignData() {
  const container = document.getElementById('sp-campaigns-grid');
  if (!container) return;
  setSkeleton(container, 4, 'sk-card');

  // Try structured list first; fall back to document folder
  let campaigns = [];
  try {
    campaigns = await fetchCampaigns();
  } catch (e) {
    console.warn('Campaigns list unavailable, trying document folder:', e.message);
  }

  if (campaigns.length) {
    // Rich campaign cards from SharePoint List
    const live      = campaigns.filter(c => c.status === 'live').length;
    const planning  = campaigns.filter(c => c.status === 'planning').length;
    const completed = campaigns.filter(c => c.status === 'completed').length;
    setMetric('sp-metric-live',      live);
    setMetric('sp-metric-planning',  planning);
    setMetric('sp-metric-completed', completed);

    container.innerHTML = campaigns.map(c => `
      <article class="camp-card" onclick="window.open('${escAttr(c.link)}','_blank')">
        <div class="camp-thumb tone-${colourForStatus(c.status)}">
          <span class="pill ${c.status}"><span class="status-dot ${dotColour(c.status)}"></span>${ucFirst(c.status)}</span>
        </div>
        <div class="camp-body">
          <div class="camp-cat">${escHtml(c.type)}</div>
          <h3 class="camp-name">${escHtml(c.title)}</h3>
          <div class="camp-dates">${escHtml(c.dates)} · ${escHtml(c.region)}</div>
          <div class="camp-kpis">
            <div><div class="k">${c.budget}</div><div class="l">Budget</div></div>
            <div><div class="k">${c.channels}</div><div class="l">Channels</div></div>
          </div>
        </div>
      </article>
    `).join('');
  } else {
    // Fallback: folders in Shared Documents/Campaigns — one folder per campaign
    try {
      const folder = HUB_CONFIG.folders?.campaigns || 'Campaigns';
      const docs = await fetchDocuments(folder);
      const items = docs.filter(d => d.isFolder); // each sub-folder is a campaign
      if (!items.length) { container.innerHTML = '<p class="sp-error">No campaigns found in SharePoint.</p>'; return; }

      // Update metrics based on folder count (all shown as active)
      setMetric('sp-metric-live',      items.length);
      setMetric('sp-metric-planning',  0);
      setMetric('sp-metric-completed', 0);

      // Cycle through accent colours for visual variety
      const tones = ['red','amber','grey'];
      container.innerHTML = items.map((d, i) => `
        <article class="camp-card" onclick="window.open('${escAttr(d.url)}','_blank')">
          <div class="camp-thumb tone-${tones[i % tones.length]}">
            <span class="pill live"><span class="status-dot green"></span>Active</span>
          </div>
          <div class="camp-body">
            <div class="camp-cat">Campaign</div>
            <h3 class="camp-name">${escHtml(d.name)}</h3>
            <div class="camp-dates">Updated ${escHtml(d.modified)}</div>
            <div class="camp-kpis">
              <div><div class="k">—</div><div class="l">Budget</div></div>
              <div><div class="k">Open folder</div><div class="l">Documents</div></div>
            </div>
          </div>
        </article>
      `).join('');
    } catch (e2) {
      container.innerHTML = `<p class="sp-error">Couldn't load campaigns: ${escHtml(e2.message)}</p>`;
    }
  }
}

// ─── Events Page Data ─────────────────────────────────────────

async function loadEventsData() {
  const container = document.getElementById('sp-events-list');
  if (!container) return;
  setSkeleton(container, 5);

  // Try structured list first; fall back to document folder
  let events = [];
  try {
    events = await fetchEvents();
  } catch (e) {
    console.warn('Events list unavailable, trying document folder:', e.message);
  }

  if (events.length) {
    // Rich list items from SharePoint List
    container.innerHTML = events.map(e => `
      <div class="sp-list-item">
        <div>
          <div class="sp-list-title">${escHtml(e.title)}</div>
          <div class="sp-list-sub">${escHtml(e.date)} · ${escHtml(e.location)}</div>
        </div>
        <span class="sp-badge ${statusBadgeClass(e.status)}">${ucFirst(e.status)}</span>
      </div>
    `).join('');
  } else {
    // Fallback: files/folders from the Events document folder
    try {
      const folder = HUB_CONFIG.folders?.events || 'Events';
      const docs = await fetchDocuments(folder);
      if (!docs.length) { container.innerHTML = '<p class="sp-error">No events found in SharePoint.</p>'; return; }
      container.innerHTML = docs.map(d => `
        <a class="sp-list-item" href="${escAttr(d.url)}" target="_blank" rel="noopener" style="text-decoration:none;color:inherit">
          <div>
            <div class="sp-list-title">${d.isFolder ? folderIcon() + ' ' : fileIcon(d.ext) + ' '}${escHtml(d.name)}</div>
            <div class="sp-list-sub">Updated ${escHtml(d.modified)}</div>
          </div>
          <span class="sp-badge upcoming">View</span>
        </a>
      `).join('');
    } catch (e2) {
      container.innerHTML = `<p class="sp-error">Couldn't load events: ${escHtml(e2.message)}</p>`;
    }
  }
}

// ─── Resources Page Data ──────────────────────────────────────

async function loadResourcesData() {
  const container = document.getElementById('sp-documents-grid');
  if (!container) return;
  setSkeleton(container, 6);
  try {
    // Fetch root of document library — shows Brand, Campaigns, Events, Launches, Products, Reports folders
    const docs = await fetchDocuments();
    if (!docs.length) { container.innerHTML = '<p class="sp-error">No files found in the document library.</p>'; return; }

    container.innerHTML = `<div class="sp-doc-grid">${
      docs.map(d => `
        <a class="sp-doc-card" href="${escAttr(d.url)}" target="_blank" rel="noopener">
          <div class="sp-doc-icon">${d.isFolder ? folderIcon() : fileIcon(d.ext)}</div>
          <div>
            <div class="sp-doc-name">${escHtml(d.name)}</div>
            <div class="sp-doc-meta">${escHtml(d.modified)}${d.size ? ' · ' + escHtml(d.size) : ''}</div>
          </div>
        </a>
      `).join('')
    }</div>`;
  } catch (e) {
    container.innerHTML = `<p class="sp-error">Couldn't load documents: ${escHtml(e.message)}</p>`;
  }
}

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

function answerPoll(btn) {
  btn.parentElement.querySelectorAll('.poll-opt').forEach(o => o.classList.remove('selected'));
  btn.classList.add('selected');
  showToast('Thanks — your vote was recorded');
}

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

function fileIcon(ext) {
  return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="18" height="18"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14,2 14,8 20,8"/></svg>`;
}

function folderIcon() {
  return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" width="18" height="18"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>`;
}

// ─── Countdown helpers ────────────────────────────────────────

function daysUntilDisplay(dateStr) {
  const d = daysUntil(dateStr);
  return d !== null ? d : '—';
}

// ─── Init ────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
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
