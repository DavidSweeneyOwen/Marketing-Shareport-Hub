

// ── WordPress News (no auth required) ────────────────────────

async function fetchWordPressNews() {
  const { apiUrl, postsPerPage } = HUB_CONFIG.wordpress;
  const url = `${apiUrl}/posts?per_page=${postsPerPage}&_fields=id,title,excerpt,date,link,jetpack_featured_media_url,_links&_embed=wp:featuredmedia`;

  const res = await fetch(url);
  if (!res.ok) throw new Error(`WordPress API returned ${res.status}`);
  const posts = await res.json();

  return posts.map(post => {
    // Try to get featured image from embed or jetpack field
    let image = post.jetpack_featured_media_url || null;
    if (!image && post._embedded?.['wp:featuredmedia']?.[0]?.source_url) {
      image = post._embedded['wp:featuredmedia'][0].source_url;
    }

    // Strip HTML tags from excerpt
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
}
