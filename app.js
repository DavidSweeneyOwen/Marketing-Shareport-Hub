

// ── WordPress News Rendering ──────────────────────────────────

async function loadWordPressNews() {
  const container = document.getElementById('sp-wp-news');
  if (!container) return;

  // Show skeleton while loading
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

// ── Intranet Homepage ─────────────────────────────────────────

async function loadHomePage() {
  await Promise.all([
    loadHeroNews(),
    loadProductNewsList(),
    startCountdown(),
  ]);
}

// Hero news grid — top 4 WordPress posts
async function loadHeroNews() {
  try {
    const posts = await fetchWordPressNews();
    if (!posts.length) return;

    // Feature article (first post)
    const feature = posts[0];
    const featureEl = document.getElementById('hero-feature');
    if (featureEl) {
      featureEl.href = feature.link;
      const img = document.getElementById('hero-feature-img');
      if (img && feature.image) img.style.backgroundImage = `url('${feature.image}')`;
      const title = document.getElementById('hero-feature-title');
      if (title) title.textContent = feature.title;
      const date = document.getElementById('hero-feature-date');
      if (date) date.textContent = feature.date;
      featureEl.classList.remove('skeleton-card');
    }

    // Side articles (posts 1-3)
    const sideContainer = document.getElementById('hero-side-articles');
    if (sideContainer && posts.length > 1) {
      sideContainer.innerHTML = posts.slice(1, 4).map(post => `
        <a class="inh-hero-side-item" href="${post.link}" target="_blank" rel="noopener">
          <div class="inh-hero-img" ${post.image ? `style="background-image:url('${post.image}')"` : 'style="background:#2A2A2A"'}></div>
          <div class="inh-hero-overlay">
            <span class="inh-tag blue" style="font-size:9px">News</span>
            <div class="inh-hero-title">${escHtml(post.title)}</div>
            <div class="inh-hero-date">${post.date}</div>
          </div>
        </a>
      `).join('');
    }
  } catch(e) {
    console.warn('Hero news load failed:', e.message);
  }
}

// Product news list — static for now, will connect to SharePoint launches
function loadProductNewsList() {
  const el = document.getElementById('home-product-news');
  if (!el) return;

  const items = [
    { tag: 'New Launch', name: 'ProGuard 6kg CO₂ Extinguisher', date: 'Launching 2 June 2026', onclick: "showPage('launches',1)" },
    { tag: 'Coming Soon', name: 'FireShield Premium Kit', date: 'Launching 19 June 2026', onclick: "showPage('launches',1)" },
    { tag: 'Back in Stock', name: 'Commander Stand', date: 'Available now', onclick: "showPage('launches',1)" },
  ];

  el.innerHTML = items.map(item => `
    <div class="inh-product-item" onclick="${item.onclick}">
      <div class="inh-product-img-placeholder">CF</div>
      <div class="inh-product-info">
        <div class="inh-product-tag">${item.tag}</div>
        <div class="inh-product-name">${item.name}</div>
        <div class="inh-product-date">${item.date}</div>
      </div>
    </div>
  `).join('');
}

// Countdown to next event
function startCountdown() {
  const target = new Date('2026-07-08T09:00:00');
  function tick() {
    const now  = new Date();
    const diff = target - now;
    if (diff <= 0) return;
    const days  = Math.floor(diff / 86400000);
    const hours = Math.floor((diff % 86400000) / 3600000);
    const mins  = Math.floor((diff % 3600000)  / 60000);
    const d = document.getElementById('cd-days');
    const h = document.getElementById('cd-hours');
    const m = document.getElementById('cd-mins');
    if (d) d.textContent = String(days).padStart(2,'0');
    if (h) h.textContent = String(hours).padStart(2,'0');
    if (m) m.textContent = String(mins).padStart(2,'0');
  }
  tick();
  setInterval(tick, 30000);
}

// Update nav active state when switching pages
const _origShowPage = typeof showPage === 'function' ? showPage : null;
function updateNavActive(id) {
  document.querySelectorAll('.inh-nav-link[id]').forEach(el => el.classList.remove('active'));
  const map = { home: 'navl-home', launches: 'navl-launches', campaigns: 'navl-campaigns', trade: 'navl-trade', training: 'navl-training' };
  const target = document.getElementById(map[id]);
  if (target) target.classList.add('active');
}
