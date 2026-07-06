/**
 * CheckFire Marketing Hub — Configuration
 *
 * NOTE ON SECRETS: nothing in this file is a secret.
 *  - tenantId / clientId are public Azure "app registration" IDs. They are
 *    designed to be visible in browser code; the actual sign-in secret never
 *    leaves Microsoft. Safe to commit.
 *  - The hub uses NO Jotform API key. The booking form is a public iframe,
 *    and the showroom calendar reads bookings from a SharePoint list using
 *    the signed-in user's token (same secure method as the other lists).
 *  - There is therefore nothing here that needs hiding in GitHub Secrets.
 */

const HUB_CONFIG = {

  tenantId:   'a865e107-f1f1-4c83-b773-130146f1deff',
  clientId:   '0cae2c21-b578-446e-8518-0855ad12d494',

  // Internal marketing-updates site (Launches / Campaigns / Events lists)
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

// ── Showroom Bookings ─────────────────────────────────────────
// The hub reads "who's coming in" from this SharePoint list on the
// MarketingHub site. Bookings land here automatically via a Power
// Automate flow when someone submits the Jotform booking form.
// See SHOWROOM-CALENDAR-SETUP.md for the list columns + flow steps.
HUB_CONFIG.showroom = {
  list: 'Showroom Bookings',
};

// ── WordPress News Feed (public, no key) ──────────────────────
HUB_CONFIG.wordpress = {
  apiUrl: 'https://www.checkfire.co.uk/wp-json/wp/v2',
  postsPerPage: 6,
};

// ── Jotform — Showroom Booking form ───────────────────────────
// Public booking form embedded as an iframe. NO API key is used or
// stored anywhere in this app — see the security note in jotform.js.
HUB_CONFIG.jotform = {
  formId: '240422414566047',
};
