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

// ── Navigation ────────────────────────────────────────────────

let currentPage = 'home';
const loadedPages = new Set();

function showPage(name) {
  document.querySelectorAll('main.page').forEach(p => p.classList.remove('active'));
  const target = document.getElementById('page-' + name);
  if (target) target.classList.add('active');
  currentPage = name;
  window.scrollTo(0, 0);

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
  // WordPress news is loaded separately (no auth needed)
  // Nothing extra needed on home for now
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
        ${item.Status    ? `<span class="sp-badge">${escHtml(item.Status)}</span>` : ''}
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
        ${item.Status    ? `<span class="sp-badge">${escHtml(item.Status)}</span>` : ''}
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
          ${dateStr      ? `<span class="sp-badge">${escHtml(dateStr)}</span>` : ''}
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

  // Skeleton while loading
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

  // WordPress news loads immediately — no sign-in needed
  loadWordPressNews();

  // Set up sign-in button href ONLY when not returning from Microsoft
  // (returning pages have ?code= in the URL — don't overwrite stored state)
  if (!window.location.search.includes('code=') && !window.location.search.includes('error=')) {
    if (typeof setupSignInButton === 'function') setupSignInButton();
  }

  // Run auth — handles redirect, refresh, or shows overlay
  const signedIn = await initAuth();

  if (signedIn) {
    // Load data for the home page on first sign-in
    loadedPages.add('home');
    loadHomeData();
  }

});
