/**
 * CheckFire Marketing Hub — Configuration
 * ─────────────────────────────────────────
 * Fill in these values once during setup.
 * See DEPLOY.md for step-by-step instructions.
 */

const HUB_CONFIG = {

  // ── Azure AD ─────────────────────────────────────────────────
  // From your App Registration in Azure Portal > Azure Active Directory
  tenantId:   'YOUR_TENANT_ID',   // e.g. 'a1b2c3d4-...'
  clientId:   'YOUR_CLIENT_ID',   // e.g. 'e5f6g7h8-...'

  // ── SharePoint ───────────────────────────────────────────────
  // The URL of your SharePoint marketing site (no trailing slash)
  sharepointSite: 'https://checkfire.sharepoint.com/sites/Marketing',

  // ── SharePoint List Names ────────────────────────────────────
  // These must match the exact list names in your SharePoint site.
  lists: {
    launches:  'Product Launches',
    campaigns: 'Campaigns',
    events:    'Events',
  },

  // ── Document Library ─────────────────────────────────────────
  documentsLibrary: 'Marketing Assets',

  // ── Redirect URI (leave blank to auto-detect) ────────────────
  redirectUri: '',

};
