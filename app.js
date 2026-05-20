/**
 * CheckFire Marketing Hub — App bootstrap, navigation & data rendering
 */

// ── Utilities ─────────────────────────────────────────────────

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Global nav injection ──────────────────────────────────────

function injectGlobalNav() {
  // Inject CSS
  const style = document.createElement('style');
  style.textContent = `
    #global-nav {
      position: fixed;
      top: 0; left: 0; right: 0;
      height: 52px;
      background: #fff;
      border-bottom: 1px solid #e5e7eb;
      display: flex;
      align-items: center;
      padding: 0 20px;
      gap: 8px;
      z-index: 9000;
      font-family: 'Manrope', sans-serif;
      box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    #global-nav .gnav-brand {
      display: flex;
      align-items: center;
      gap: 8px;
      font-weight: 700;
      font-size: 15px;
      color: #111;
      cursor: pointer;
      text-decoration: none;
      flex-shrink: 0;
    }
    #global-nav .gnav-brand span { color: #D72B2B; }
    #global-nav .gnav-links {
      display: flex;
      align-items: center;
      gap: 2px;
      margin-left: 16px;
      flex: 1;
    }
    #global-nav .gnav-link {
      padding: 6px 12px;
      border-radius: 6px;
      font-size: 13px;
      font-weight: 500;
      color: #444;
      cursor: pointer;
      text-decoration: none;
      white-space: nowrap;
      transition: background 0.15s, color 0.15s;
    }
    #global-nav .gnav-link:hover,
    #global-nav .gnav-link.active { background: #f3f4f6; color: #111; }
    #global-nav .gnav-user {
      display: none;
      align-items: center;
      gap: 8px;
      margin-left: auto;
      flex-shrink: 0;
    }
    #global-nav .gnav-avatar {
      width: 30px; height: 30px;
      border-radius: 50%;
      background: #D72B2B;
      color: #fff;
      font-size: 11px;
      font-weight: 700;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    #global-nav .gnav-name {
      font-size: 13px;
      font-weight: 600;
      color: #111;
    }
    #global-nav .gnav-signout {
      font-size: 12px;
      color: #888;
      background: none;
      border: none;
      cursor: pointer;
      padding: 4px 8px;
      border-radius: 4px;
    }
    #global-nav .gnav-signout:hover { background: #f3f4f6; color: #333; }
    /* Push all pages down so nav doesn't cover content */
    main.page { padding-top: 52px !important; }
  `;
  document.head.appendChild(style);

  // Inject nav HTML
  const nav = document.createElement('nav');
  nav.id = 'global-nav';
  nav.innerHTML = `
    <div class="gnav-brand" onclick="showPage('home')">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none">
        <path d="M12 2C8 7 6 10 6 14a6 6 0 0 0 12 0c0-4-2-7-6-12z" fill="#D72B2B"/>
        <path d="M12 10c-1 2.5-2 4-2 5.5a2 2 0 0 0 4 0c0-1.5-1-3-2-5.5z" fill="#fff" opacity="0.7"/>
      </svg>
      Check<span>Fire</span> Marketing Hub
    </div>
    <div class="gnav-links">
      <a class="gnav-link active" id="gnav-home"      onclick="showPage('home')">Home</a>
      <a class="gnav-link"       id="gnav-launches"   onclick="showPage('launches')">Launches</a>
      <a class="gnav-link"       id="gnav-campaigns"  onclick="showPage('campaigns')">Campaigns</a>
      <a class="gnav-link"       id="gnav-trade"      onclick="showPage('trade')">Trade &amp; Events</a>
      <a class="gnav-link"       id="gnav-training"   onclick="showPage('training')">Resources</a>
    </div>
    <div id="nav-user-info" class="gnav-user">
      <div id="nav-user-avatar" class="gnav-avatar">?</div>
      <span id="nav-user-name" class="gnav-name">User</span>
      <button class="gnav-signout" onclick="signOut()">Sign out</button>
    </div>
  `;
  document.body.insertBefore(nav, document.body.firstChild);
}

// ── Navigation ────────────────────────────────────────────────

let currentPage = 'home';
const loadedPages = new Set();

// Accepts optional second arg for backward-compat with onclick="showPage('x',2)"
function showPage(name, _idx) {
  document.querySelectorAll('main.page').forEach(p => p.classList.remove('active'));
  const target = document.getElementById('page-' + name);
  if (target) target.classList.add('active');
  currentPage = name;
  window.scrollTo(0, 0);

  // Update active nav link
  document.querySelectorAll('#global-nav .gnav-link').forEach(l => l.classList.remove('active'));
  const activeLink = document.getElementById('gnav-' + name);
  if (activeLink) activeLink.classList.add('active');

  // Load data for this page if signed in and not already loaded
  if (AUTH.token && !loadedPages.has(name)) {
    loadedPages.add(name);
    loadPageData(name);
  }
}

async function loadPageData(name) {
  switch (name) {
    case 'home':      await loadHomeData();      break;
    case 'launches':  await loadLaunchData();    break;
    case 'campaigns': await loadCampaignData();  break;
    case 'trade':     await loadEventsData();    break;
    case 'training':  await loadResourcesData(); break;
  }
}

// ── Data loaders ──────────────────────────────────────────────

async function loadHomeData() {
  // WordPress news loaded separately (no auth required)
}

async function loadLaunchData() {
  const el = document.getElementById('sp-launches-list');
  if (!el) return;
  el.innerHTML = '<p style="color:#888;font-size:14px;padding:16px 0">Loading launches from SharePoint…</p>';
  try {
    const items = await fetchLaunches();
    if (!items.length) {
      el.innerHTML = '<p style="color:#888;font-size:14px;padding:16px 0">No launches found in SharePoint.</p>';
      return;
    }
    el.innerHTML = items.map(item => `
      <div class="sp-card">
        <div class="sp-card-title">${escHtml(item.Title || 'Untitled')}</div>
        ${item.Status      ? `<span class="sp-badge">${escHtml(item.Status)}</span>` : ''}
        ${item.Description ? `<p class="sp-card-desc">${escHtml(item.Description)}</p>` : ''}
      </div>`).join('');
  } catch (e) {
    el.innerHTML = `<p class="sp-error">${escHtml(e.message)}</p>`;
  }
}

async function loadCampaignData() {
  const el = document.getElementById('sp-campaigns-grid');
  if (!el) return;
  el.innerHTML = '<p style="color:#888;font-size:14px;padding:16px 0">Loading campaigns from SharePoint…</p>';
  try {
    const items = await fetchCampaigns();
    if (!items.length) {
      el.innerHTML = '<p style="color:#888;font-size:14px;padding:16px 0">No campaigns found in SharePoint.</p>';
      return;
    }
    el.innerHTML = items.map(item => `
      <div class="sp-card">
        <div class="sp-card-title">${escHtml(item.Title || 'Untitled')}</div>
        ${item.Status      ? `<span class="sp-badge">${escHtml(item.Status)}</span>` : ''}
        ${item.Description ? `<p class="sp-card-desc">${escHtml(item.Description)}</p>` : ''}
      </div>`).join('');
  } catch (e) {
    el.innerHTML = `<p class="sp-error">${escHtml(e.message)}</p>`;
  }
}

async function loadEventsData() {
  const el = document.getElementById('sp-events-list');
  if (!el) return;
  el.innerHTML = '<p style="color:#888;font-size:14px;padding:16px 0">Loading events from SharePoint…</p>';
  try {
    const items = await fetchEvents();
    if (!items.length) {
      el.innerHTML = '<p style="color:#888;font-size:14px;padding:16px 0">No events found in SharePoint.</p>';
      return;
    }
    el.innerHTML = items.map(item => {
      const dateStr = item.EventDate
        ? new Date(item.EventDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })
        : '';
      return `
        <div class="sp-card">
          <div class="sp-card-title">${escHtml(item.Title || 'Untitled')}</div>
          ${dateStr       ? `<span class="sp-badge">${escHtml(dateStr)}</span>` : ''}
          ${item.Location ? `<p class="sp-card-desc">📍 ${escHtml(item.Location)}</p>` : ''}
        </div>`;
    }).join('');
  } catch (e) {
    el.innerHTML = `<p class="sp-error">${escHtml(e.message)}</p>`;
  }
}

async function loadResourcesData() {
  const el = document.getElementById('sp-documents-grid');
  if (!el) return;
  el.innerHTML = '<p style="color:#888;font-size:14px;padding:16px 0">Loading documents from SharePoint…</p>';
  try {
    const items = await fetchDocuments();
    if (!items.length) {
      el.innerHTML = '<p style="color:#888;font-size:14px;padding:16px 0">No documents found in SharePoint.</p>';
      return;
    }
    el.innerHTML = items.map(item => {
      const updated = item.lastModifiedDateTime
        ? new Date(item.lastModifiedDateTime).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })
        : '';
      return `
        <a class="sp-card sp-doc-card" href="${item.webUrl || '#'}" target="_blank" rel="noopener">
          <div class="sp-doc-icon">📄</div>
          <div class="sp-card-title">${escHtml(item.name || 'Document')}</div>
          ${updated ? `<div class="sp-card-meta">Updated ${escHtml(updated)}</div>` : ''}
        </a>`;
    }).join('');
  } catch (e) {
    el.innerHTML = `<p class="sp-error">${escHtml(e.message)}</p>`;
  }
}

// ── WordPress News ────────────────────────────────────────────

async function loadWordPressNews() {
  const container = document.getElementById('sp-wp-news');
  if (!container) return;

  container.innerHTML = `
    <div class="wp-news-grid">
      ${Array(3).fill('<div class="wp-news-card skeleton" style="height:200px;border-radius:12px"></div>').join('')}
    </div>`;

  try {
    const posts = await fetchWordPressNews();
    if (!posts.length) {
      container.innerHTML = '<p class="sp-error">No news posts found.</p>';
      return;
    }
    container.innerHTML = `<div class="wp-news-grid">${posts.map(post => `
      <a class="wp-news-card" href="${post.link}" target="_blank" rel="noopener">
        ${post.image
          ? `<div class="wp-news-img" style="background-image:url('${post.image}')"></div>`
          : `<div class="wp-news-img wp-news-img-placeholder"></div>`}
        <div class="wp-news-body">
          <div class="wp-news-date">${post.date}</div>
          <h3 class="wp-news-title">${escHtml(post.title)}</h3>
          <p class="wp-news-excerpt">${escHtml(post.excerpt)}</p>
        </div>
      </a>`).join('')}
    </div>`;
  } catch (e) {
    container.innerHTML = `<p class="sp-error">Couldn't load news: ${escHtml(e.message)}</p>`;
  }
}

// ── Bootstrap ─────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', async () => {

  // Inject persistent top navigation bar
  injectGlobalNav();

  // Load WordPress news immediately — no auth needed
  loadWordPressNews();

  // Set up sign-in button ONLY when not returning from Microsoft
  if (!window.location.search.includes('code=') && !window.location.search.includes('error=')) {
    if (typeof setupSignInButton === 'function') setupSignInButton();
  }

  // Run auth flow
  const signedIn = await initAuth();

  if (signedIn) {
    loadedPages.add('home');
    loadHomeData();
  }

});
