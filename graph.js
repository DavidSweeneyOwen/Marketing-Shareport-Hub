/**
 * CheckFire Marketing Hub — Microsoft Graph API layer
 * ─────────────────────────────────────────────────────
 * All SharePoint/Graph calls go through here.
 * Each function returns structured data ready for the UI to render.
 */

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// ── Core fetch wrapper ────────────────────────────────────────

async function graphFetch(path, params = '') {
  const token = await getAccessToken();
  if (!token) throw new Error('No access token');

  const url = `${GRAPH_BASE}${path}${params ? '?' + params : ''}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `HTTP ${res.status}`);
  }
  return res.json();
}

// ── Site resolution ───────────────────────────────────────────

let _siteId = null;
async function getSiteId() {
  if (_siteId) return _siteId;
  // Convert site URL to Graph site ID
  const url = new URL(HUB_CONFIG.sharepointSite);
  const host = url.hostname;                         // checkfire.sharepoint.com
  const path = url.pathname.replace(/^\//, '');      // sites/Marketing
  const data = await graphFetch(`/sites/${host}:/${path}`);
  _siteId = data.id;
  return _siteId;
}

async function getListId(listName) {
  const siteId = await getSiteId();
  const data = await graphFetch(`/sites/${siteId}/lists`, `$filter=displayName eq '${encodeURIComponent(listName)}'&$select=id,displayName`);
  const list = data.value?.[0];
  if (!list) throw new Error(`List "${listName}" not found`);
  return list.id;
}

// ── Launches ──────────────────────────────────────────────────

async function fetchLaunches() {
  const siteId = await getSiteId();
  const listId = await getListId(HUB_CONFIG.lists.launches);
  const data = await graphFetch(
    `/sites/${siteId}/lists/${listId}/items`,
    `$expand=fields&$select=fields&$orderby=fields/LaunchDate asc&$top=10`
  );
  return (data.value || []).map(item => {
    const f = item.fields;
    return {
      title:       f.Title        || 'Untitled',
      sku:         f.SKU          || '—',
      launchDate:  f.LaunchDate   || null,
      status:      f.Status       || 'planning',
      description: f.Description  || '',
      price:       f.RRP          || '',
      link:        f.LinkURL      || '#',
    };
  });
}

// ── Campaigns ─────────────────────────────────────────────────

async function fetchCampaigns() {
  const siteId = await getSiteId();
  const listId = await getListId(HUB_CONFIG.lists.campaigns);
  const data = await graphFetch(
    `/sites/${siteId}/lists/${listId}/items`,
    `$expand=fields&$select=fields&$orderby=fields/StartDate desc&$top=20`
  );
  return (data.value || []).map(item => {
    const f = item.fields;
    const start = f.StartDate ? new Date(f.StartDate).toLocaleDateString('en-GB', { day:'numeric', month:'short' }) : '';
    const end   = f.EndDate   ? new Date(f.EndDate).toLocaleDateString('en-GB', { day:'numeric', month:'short' }) : '';
    return {
      title:    f.Title       || 'Untitled',
      type:     f.CampaignType || 'Campaign',
      status:   (f.Status     || 'planning').toLowerCase(),
      dates:    start && end ? `${start} – ${end}` : start || '—',
      budget:   f.Budget      ? `£${Number(f.Budget).toLocaleString('en-GB')}` : '—',
      channels: f.Channels    || '—',
      region:   f.Region      || 'UK',
      link:     f.LinkURL     || '#',
    };
  });
}

// ── Events ────────────────────────────────────────────────────

async function fetchEvents() {
  const siteId = await getSiteId();
  const listId = await getListId(HUB_CONFIG.lists.events);
  const data = await graphFetch(
    `/sites/${siteId}/lists/${listId}/items`,
    `$expand=fields&$select=fields&$orderby=fields/EventDate asc&$top=20`
  );
  return (data.value || []).map(item => {
    const f = item.fields;
    const date = f.EventDate ? new Date(f.EventDate) : null;
    return {
      title:    f.Title       || 'Untitled',
      date:     date ? date.toLocaleDateString('en-GB', { day:'numeric', month:'short', year:'numeric' }) : '—',
      dateObj:  date,
      location: f.Location    || '—',
      type:     f.EventType   || 'Event',
      status:   f.Status      || 'upcoming',
      link:     f.LinkURL     || '#',
    };
  });
}

// ── Documents (Resources) ─────────────────────────────────────

async function fetchDocuments(folderPath = '') {
  const siteId = await getSiteId();
  // Find the drive (document library)
  const drives = await graphFetch(`/sites/${siteId}/drives`, `$select=id,name`);
  const drive = drives.value?.find(d => d.name === HUB_CONFIG.documentsLibrary)
             || drives.value?.[0];
  if (!drive) throw new Error('Document library not found');

  const path = folderPath ? `/root:/${folderPath}:/children` : '/root/children';
  const data = await graphFetch(`/drives/${drive.id}${path}`, `$select=id,name,lastModifiedDateTime,size,webUrl,file,folder&$orderby=lastModifiedDateTime desc&$top=30`);
  return (data.value || []).map(item => ({
    id:       item.id,
    name:     item.name,
    isFolder: !!item.folder,
    modified: item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).toLocaleDateString('en-GB', { day:'numeric', month:'short', year:'numeric' }) : '—',
    size:     item.size ? formatFileSize(item.size) : null,
    url:      item.webUrl || '#',
    ext:      item.name.split('.').pop().toUpperCase(),
  }));
}

// ── Recent activity (home page) ───────────────────────────────

async function fetchRecentFiles() {
  // Uses /me/drive/recent — shows files the signed-in user recently accessed
  const data = await graphFetch('/me/drive/recent', `$select=id,name,lastModifiedDateTime,webUrl,remoteItem&$top=5`);
  return (data.value || []).map(item => ({
    name:     item.name,
    modified: item.lastModifiedDateTime ? timeAgo(new Date(item.lastModifiedDateTime)) : '',
    url:      item.webUrl || item.remoteItem?.webUrl || '#',
  }));
}

// ── Utilities ─────────────────────────────────────────────────

function formatFileSize(bytes) {
  if (bytes < 1024)       return bytes + ' B';
  if (bytes < 1048576)    return (bytes / 1024).toFixed(0) + ' KB';
  if (bytes < 1073741824) return (bytes / 1048576).toFixed(1) + ' MB';
  return (bytes / 1073741824).toFixed(1) + ' GB';
}

function timeAgo(date) {
  const diff = Math.floor((Date.now() - date) / 1000);
  if (diff < 60)    return 'just now';
  if (diff < 3600)  return Math.floor(diff/60) + 'm ago';
  if (diff < 86400) return Math.floor(diff/3600) + 'h ago';
  return Math.floor(diff/86400) + 'd ago';
}

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const now = new Date(); now.setHours(0,0,0,0);
  const target = new Date(dateStr);
  return Math.max(0, Math.round((target - now) / 86400000));
}

function statusBadgeClass(status) {
  const s = (status || '').toLowerCase();
  if (s === 'live' || s === 'active' || s === 'launched') return 'live';
  if (s === 'planning' || s === 'draft' || s === 'pending') return 'planning';
  if (s === 'completed' || s === 'done' || s === 'closed') return 'completed';
  return 'upcoming';
}
