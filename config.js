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

  // Product Portal — a second SharePoint site whose library is browsed
  // and previewed INSIDE the hub (it never bounces out to SharePoint).
  // Reachable with the same delegated Sites.Read.All / Files.Read.All
  // scopes the hub already uses — no new admin consent needed as long as
  // the signed-in user has access to the site.
  productPortalSite: 'https://checkfireltd.sharepoint.com/sites/CheckFireProductPortal',

  // Libraries shown on the Product Portal tab (first = default). The
  // portal keeps its datasheets in a separate "Data Sheets" library,
  // so the in-hub browser offers both.
  productPortalLibraries: ['Data Sheets', 'Documents'],

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

  // NOTE: campaign/launch asset blocks are no longer configured here.
  // The detail pages discover the item's real sub-folders in SharePoint
  // (Documents/Campaigns/<name>/... or Documents/Launches/<name>/...)
  // and show whatever exists - marketing can add/rename folders freely.

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

// ── Videos — pulled onto the hub home page ────────────────────
// WordPress: public media API on checkfire.co.uk (video uploads).
// SharePoint: the Media Portal site's library is searched for video
// files (e.g. the "03. Videos" folder) using the signed-in user's
// token — same read-only Graph access as the rest of the hub.
HUB_CONFIG.videos = {
  mediaPortalSite: 'https://checkfireltd.sharepoint.com/sites/CheckFireMediaPortal',
  max: 6,            // how many to show on the home page
  maxAgeMonths: 3,   // only show videos uploaded within the last N months (0 = no limit)
};

// ── WordPress News Feed (public, no key) ──────────────────────
// posts  → "Latest blogs" carousel on the home page.
// pages  → "Updated landing pages" carousel (public WordPress pages,
//          newest-modified first). Set landingPageParent to a page ID
//          to only show its children, or leave 0 for all pages.
HUB_CONFIG.wordpress = {
  apiUrl: 'https://www.checkfire.co.uk/wp-json/wp/v2',
  postsPerPage: 8,
  pagesPerPage: 8,
  landingPageParent: 0,
};

// ── Showroom / marketing calendar ─────────────────────────────
// The home-page calendar marks three kinds of date: showroom visits
// (red), product launches (blue) and campaign runs (amber). Users can
// "Subscribe" to it. On a static site the Subscribe button downloads an
// .ics containing every marked date, which imports into Outlook/Google.
// If you later publish a live SharePoint/Outlook calendar, paste its
// webcal:// or https .ics feed URL below and Subscribe will use that
// instead (a true auto-updating subscription).
HUB_CONFIG.calendar = {
  feedUrl: '',   // e.g. 'webcal://outlook.office365.com/owa/calendar/.../reachcalendar.ics'
};

// ── Jotform — Showroom Booking form ───────────────────────────
// Public booking form embedded as an iframe. NO API key is used or
// stored anywhere in this app — see the security note in jotform.js.
HUB_CONFIG.jotform = {
  formId: '240422414566047',

  // Azure Function proxy that holds the Jotform API key server-side.
  // Leave '' and the hub reads bookings from the SharePoint list only.
  // When set (e.g. 'https://checkfire-jotform.azurewebsites.net/api/bookings'),
  // the hub also pulls live submissions straight from Jotform via the proxy.
  // The key NEVER ships to the browser — see azure-function/README.md.
  proxyUrl: '',
};

// ── Social — LinkedIn + in-house comms feed (home page) ───────
// LinkedIn has no free auto-updating page feed, so we embed specific
// posts. In LinkedIn: open a post ▸ ••• ▸ "Embed this post" ▸ copy the
// src="..." from the <iframe> ▸ paste the URLs below (newest first).
// The in-house comms feed is a Twitter-style stream driven by a
// SharePoint list on the MarketingHub site (see the setup guide for
// the list columns). Both panels hide themselves when empty.
HUB_CONFIG.social = {
  linkedInPageUrl: 'https://www.linkedin.com/company/checkfire/',
  linkedInEmbeds: [
    // 'https://www.linkedin.com/embed/feed/update/urn:li:share:0000000000000000000',
  ],
  commsList: 'Comms',   // SharePoint list of in-house announcements
  commsMax: 8,
};
