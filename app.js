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
    case 'launches':  await loadLaunchData();    break;
    case 'campaigns': await loadCampaignData();  break;
    case 'trade':     await loadEventsData();    break;
    case 'training':  await loadResourcesData(); break;
  }
}

// ─── Home Page Data ───────────────────────────────────────────

async function loadHomeData() {
  try {
    const files = await fetchRecentFiles();
    renderRecentFiles(files);
  } catch (e) {
    console.warn('Recent files unavailable:', e.message);
  }
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
  try {
    const launches = await fetchLaunches();
    if (!launches.length) { container.innerHTML = '<p class="sp-error">No launches found in SharePoint.</p>'; return; }

    // Featured (first upcoming/live launch goes in the hero)
    const featured = launches.find(l => l.status.toLowerCase() !== 'completed') || launches[0];
    updateLaunchHero(featured);

    // List of all launches
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
  } catch (e) {
    container.innerHTML = `<p class="sp-error">Couldn't load launches: ${escHtml(e.message)}</p>`;
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
  try {
    const campaigns = await fetchCampaigns();
    if (!campaigns.length) { container.innerHTML = '<p class="sp-error">No campaigns found in SharePoint.</p>'; return; }

    // Update metric counts
    const live      = campaigns.filter(c => c.status === 'live').length;
    const planning  = campaigns.filter(c => c.status === 'planning').length;
    const completed = campaigns.filter(c => c.status === 'completed').length;
    setMetric('sp-metric-live',      live);
    setMetric('sp-metric-planning',  planning);
    setMetric('sp-metric-completed', completed);

    container.innerHTML = campaigns.map(c => `
      <article class="camp-card" onclick="window.open('${c.link}','_blank')">
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
  } catch (e) {
    container.innerHTML = `<p class="sp-error">Couldn't load campaigns: ${escHtml(e.message)}</p>`;
  }
}

// ─── Events Page Data ─────────────────────────────────────────

async function loadEventsData() {
  const container = document.getElementById('sp-events-list');
  if (!container) return;
  setSkeleton(container, 5);
  try {
    const events = await fetchEvents();
    if (!events.length) { container.innerHTML = '<p class="sp-error">No events found in SharePoint.</p>'; return; }

    container.innerHTML = events.map(e => `
      <div class="sp-list-item">
        <div>
          <div class="sp-list-title">${escHtml(e.title)}</div>
          <div class="sp-list-sub">${escHtml(e.date)} · ${escHtml(e.location)}</div>
        </div>
        <span class="sp-badge ${statusBadgeClass(e.status)}">${ucFirst(e.status)}</span>
      </div>
    `).join('');
  } catch (e) {
    container.innerHTML = `<p class="sp-error">Couldn't load events: ${escHtml(e.message)}</p>`;
  }
}

// ─── Resources Page Data ──────────────────────────────────────

async function loadResourcesData() {
  const container = document.getElementById('sp-documents-grid');
  if (!container) return;
  setSkeleton(container, 6);
  try {
    const docs = await fetchDocuments();
    if (!docs.length) { container.innerHTML = '<p class="sp-error">No files found in the document library.</p>'; return; }

    container.innerHTML = `<div class="sp-doc-grid">${
      docs.map(d => `
        <a class="sp-doc-card" href="${d.url}" target="_blank" rel="noopener">
          <div class="sp-doc-icon">${d.isFolder ? folderIcon() : fileIcon(d.ext)}</div>
          <div>
            <div class="sp-doc-name">${escHtml(d.name)}</div>
            <div class="sp-doc-meta">${d.modified}${d.size ? ' · ' + d.size : ''}</div>
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

// ─── Countdown helpers (preserved from prototype) ─────────────

function daysUntilDisplay(dateStr) {
  const d = daysUntil(dateStr);
  return d !== null ? d : '—';
}

// ─── Init ────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {
  // Countdowns
  const launchDays = document.getElementById('launchDays');
  if (launchDays) launchDays.textContent = '—';

  // Init auth, then load home page data
  const authed = await initAuth();
  if (authed || window.HUB_DEMO_MODE) {
    await loadPageData('home');
    animateBars();
    setTimeout(animateBars, 400);
  }
});
