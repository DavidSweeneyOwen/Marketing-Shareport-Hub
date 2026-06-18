/**
 * CheckFire Marketing Hub — Configuration
 */

const HUB_CONFIG = {

  tenantId:   'a865e107-f1f1-4c83-b773-130146f1deff',
  clientId:   '0cae2c21-b578-446e-8518-0855ad12d494',

  sharepointSite: 'https://checkfireltd.sharepoint.com/sites/MarketingHub',

  lists: {
    launches:  'Product Launches',
    campaigns: 'Campaigns',
    events:    'Events',
  },

  documentsLibrary: 'Documents',

  folders: {
    campaigns: 'Campaigns',
    launches:  'Launches',
    events:    'Events',
  },

  redirectUri: 'https://davidsweeneyowen.github.io/Marketing-Shareport-Hub/',

};


// ── WordPress News Feed ───────────────────────────────────────
// Pulls latest posts directly from the CheckFire website.
// No authentication needed — uses the public WordPress REST API.
HUB_CONFIG.wordpress = {
  apiUrl: 'https://www.checkfire.co.uk/wp-json/wp/v2',
  postsPerPage: 6,
};

// ── Jotform — Showroom Booking ────────────────────────────────
// EU API key from your Jotform account (My Account → API Keys)
HUB_CONFIG.jotform = {
  apiKey:  'YOUR_JOTFORM_API_KEY',   // ← paste your EU API key here
  formId:  '240422414566047',
  apiBase: 'https://eu-api.jotform.com',
};
