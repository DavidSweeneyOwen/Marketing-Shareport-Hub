/**
 * CheckFire Marketing Hub — Microsoft Graph & WordPress API calls
 */

// ── Core Graph fetch ──────────────────────────────────────────

async function graphFetch(path, params) {
  const token = await getAccessToken();
  if (!token) throw new Error('Not signed in');

  let url = `https://graph.microsoft.com/v1.0${path}`;
  if (params) url += '?' + new URLSearchParams(params);

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    const msg = err?.error?.message || `Graph API error ${res.status}`;
    throw new Error(msg);
  }

  return res.json();
}

// ── SharePoint site resolution ────────────────────────────────

let _siteId = null;

async function getSiteId() {
  if (_siteId) return _siteId;
  // Extract hostname and path from the SharePoint site URL
  const url = new URL(HUB_CONFIG.sharepointSite);
  const hostname = url.hostname;                     // checkfireltd.sharepoint.com
  const sitePath = url.pathname.replace(/^\//, ''); // sites/CheckFireMediaPortal
  const data = await graphFetch(`/sites/${hostname}:/${sitePath}`);
  _siteId = data.id;
  return _siteId;
}

// ── SharePoint list fetch helper ──────────────────────────────

async function fetchListItems(listName, selectFields) {
  const siteId = await getSiteId();
  const data = await graphFetch(
    `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`,
    { expand: 'fields', $select: 'fields', $top: 50 }
  );
  return (data.value || []).map(item => item.fields || {});
}

// ── SharePoint data loaders ───────────────────────────────────

async function fetchLaunches() {
  return fetchListItems(HUB_CONFIG.lists.launches);
}

async function fetchCampaigns() {
  return fetchListItems(HUB_CONFIG.lists.campaigns);
}

async function fetchEvents() {
  return fetchListItems(HUB_CONFIG.lists.events);
}

async function fetchDocuments() {
  const siteId = await getSiteId();
  const data = await graphFetch(
    `/sites/${siteId}/drive/root/children`,
    { $select: 'name,webUrl,lastModifiedDateTime,file', $top: 50 }
  );
  // Only return actual files (not folders)
  return (data.value || []).filter(item => item.file);
}

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
