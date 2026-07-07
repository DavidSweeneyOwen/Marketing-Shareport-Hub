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
    loadWordPressNews(),
    loadProductNewsList(),
    startCountdown(),
    loadHomeVideos(),
  ]);
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
  return fetchListItems(HUB_CONFIG.lists.launches).then(items => {
    if (!items.length) {
      el.innerHTML = '<p class="prose dim" style="padding:12px 0">No product launches in SharePoint yet — add items to the Product Launches list.</p>';
      return;
    }
    const sorted = [...items].sort((a, b) => String(b.LaunchDate || '').localeCompare(String(a.LaunchDate || '')));
    el.innerHTML = sorted.slice(0, 3).map(f => `
      <div class="inh-product-item" onclick="showPage('launches',1)">
        <div class="inh-product-img-placeholder">CF</div>
        <div class="inh-product-info">
          <div class="inh-product-tag">${escHtml(f.Status || 'Launch')}</div>
          <div class="inh-product-name">${escHtml(f.Title || 'Untitled')}</div>
          <div class="inh-product-date">${f.LaunchDate ? formatDate(f.LaunchDate) : ''}</div>
        </div>
      </div>
    `).join('');
  }).catch(e => {
    console.warn('Product news unavailable:', e.message);
    el.innerHTML = '<p class="prose dim" style="padding:12px 0">Sign in to see the latest product launches.</p>';
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
