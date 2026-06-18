/**
 * CheckFire Marketing Hub — Configuration
 * ─────────────────────────────────────────
 * Fill in these values once during setup.
 * See DEPLOY.md for step-by-step instructions.
 */

const HUB_CONFIG = {

  // ── Azure AD ─────────────────────────────────────────────────
  // From your App Registration in Azure Portal > Azure Active Directory
  tenantId:   'YOUR_TENANT_ID',   // e.g. 'a865e107-f1f1-4c83-b773-130146f1deff'
  clientId:   'YOUR_CLIENT_ID',   // e.g. '0cae2c21-b578-446e-8518-0855ad12d494'

  // ── SharePoint ───────────────────────────────────────────────
  // The URL of your SharePoint marketing site (no trailing slash)
  sharepointSite: 'https://checkfireltd.sharepoint.com/sites/MarketingHub',

  // ── SharePoint List Names ────────────────────────────────────
  // These must match the exact list names in your SharePoint site.
  lists: {
    launches:  'Product Launches',
    campaigns: 'Campaigns',
    events:    'Events',
  },

  // ── Document Library ─────────────────────────────────────────
  // The display name of the document library (default SharePoint library is 'Documents')
  documentsLibrary: 'Documents',

  // ── Document Library Folders ─────────────────────────────────
  // Sub-folder paths within the library for each content type.
  // Update these to match the folder names you've created in SharePoint.
  folders: {
    campaigns: 'Campaigns',   // confirmed: Shared Documents/Campaigns
    launches:  'Launches',    // update if named differently
    events:    'Events',      // update if named differently
  },

  // ── Redirect URI (leave blank to auto-detect) ────────────────
  redirectUri: '',

};
