

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
