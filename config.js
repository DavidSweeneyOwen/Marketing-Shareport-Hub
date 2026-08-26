/**
 * CheckFire Marketing Hub — Configuration
 *
 * NOTE ON SECRETS: nothing in this file is a secret.
 *  - tenantId / clientId are public Azure "app registration" IDs. They are
 *    designed to be visible in browser code; the actual sign-in secret never
 *    leaves Microsoft. Safe to commit.
 *  - The hub uses NO Jotform or YouTube API key. Both are read through the
 *    Azure Function proxy, which holds the keys server-side.
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

  // Fallback asset blocks for the Campaign / Product Launch detail pages.
  // The hub now reads the REAL sub-folders inside
  //   Documents/<Campaigns|Launches>/<item name>/
  // and builds a block for each one, with a live file count. These
  // entries are only used when that folder doesn't exist yet, so the
  // page still shows something sensible on day one.
  campaignAssetBlocks: [
    { label: 'Infographic',         folder: 'Infographic'         },
    { label: 'Email signature',     folder: 'Email signature'     },
    { label: 'Email',               folder: 'Email'               },
    { label: 'Data card',           folder: 'Data card'           },
    { label: 'Social media assets', folder: 'Social media assets' },
    { label: 'PR activities',       folder: 'PR activities'       },
  ],

  redirectUri: 'https://davidsweeneyowen.github.io/Marketing-Shareport-Hub/',

};

// ── Notices / alerts banner (home page) ───────────────────────
// Marketing post short updates here: "product delayed", "website down",
// "issue we're working on". Driven by a SharePoint list on MarketingHub.
//
// List "Notices" — columns:
//   Title    (single line)  — the headline, e.g. "Website maintenance"
//   Message  (multi-line)   — the detail
//   Type     (choice)       — Info | Warning | Alert   (colour of the bar)
//   Active   (Yes/No)       — untick to hide without deleting
//   StartDate / EndDate (date, both optional) — auto show/hide window
//   Link     (hyperlink, optional) — "More info →"
//
// The bar hides itself completely when there is nothing active.
HUB_CONFIG.notices = {
  list: 'Notices',
  max:  3,          // most recent N active notices
};

// ── Showroom Bookings ─────────────────────────────────────────
// The hub reads "who's coming in" from this SharePoint list on the
// MarketingHub site, and (when the proxy below is live) direct from
// Jotform as well. See SHOWROOM-CALENDAR-SETUP.md.
HUB_CONFIG.showroom = {
  list: 'Showroom Bookings',
};

// ── Training calendar (home page) ─────────────────────────────
// Internal and external training events — e.g. the in-house sessions
// Josh runs around a product launch, or an external course someone is
// booked on. Driven by a SharePoint list on MarketingHub.
//
// List "Training Events" — columns:
//   Title        (single line)    — what the session is
//   TrainingDate (date)           — when
//   EndDate      (date, optional) — for multi-day courses
//   Trainer      (single line, optional) — who is running it
//   Location     (single line, optional) — room, site or "Teams"
//   TrainingType (choice)         — Internal | External
//   Notes        (multi-line, optional)
//   Link         (hyperlink, optional) — joining link or booking page
//
// Training dates also appear as green markers on the marketing calendar.
HUB_CONFIG.training = {
  list: 'Training Events',
  max:  8,           // how many upcoming sessions to list
};

// ── Trade & customer events ───────────────────────────────────
// The Trade & Events page is driven entirely by the "Events" folder in
// the marketing Documents library — NOT by a SharePoint list. Each
// sub-folder is one event (e.g. "FSE 2027"), and everything the sales
// team needs in the run-up lives inside it. A four-digit year anywhere
// in the folder name splits upcoming from previous.
HUB_CONFIG.tradeEvents = {
  folder: 'Events',
};

// ── Videos — pulled onto the hub home page ────────────────────
// PRIMARY SOURCE: the CheckFire YouTube channel, read through the Azure
// Function proxy so the YouTube API key stays server-side (same pattern
// as the Jotform proxy — never put an API key in this file).
//
// HOSTNAME — verified live 4 Aug 2026. Flex Consumption apps get a
// region-scoped name with a random suffix; it is NOT
// checkfire-jotform.azurewebsites.net (that resolves to nothing, which
// is why the hub silently showed no bookings until 4 Aug). Don't
// "tidy" this back to the short form. Leave '' to disable the proxy
// and fall back to the other sources.
//
// Secondary sources are kept as a safety net:
//   WordPress  — video uploads on checkfire.co.uk
//   SharePoint — the Media Portal library ("03. Videos")
//
// CHANNEL CONFIRMED 4 Aug 2026. The handle guessed on 31 July
// (@checkfireltd) does not exist — the real channel is @CheckFireGroup.
// For the Function app's settings, use the channel ID rather than the
// handle; it survives a rename:
//   YOUTUBE_CHANNEL_ID = UC9EwvNr5cfQJW7GRrqCyphg
HUB_CONFIG.videos = {
  youtube: {
    proxyUrl:   'https://checkfire-jotform-fhagcybsfvg5fth8.uksouth-01.azurewebsites.net/api/videos',
    channelUrl: 'https://www.youtube.com/@CheckFireGroup',

    // KEYLESS FALLBACK — added 26 Aug 2026.
    // The proxy above still answers
    //   {"error":"YOUTUBE_API_KEY app setting is not configured."}
    // so the hub was showing no YouTube videos at all. It no longer has
    // to wait for that key: every YouTube channel has an "uploads"
    // playlist whose id is the channel id with UC swapped for UU, and
    // that playlist embeds with NO API key, NO quota and NO Google
    // account — and it refreshes itself the moment marketing publish.
    //
    //   channel  UC9EwvNr5cfQJW7GRrqCyphg
    //   uploads  UU9EwvNr5cfQJW7GRrqCyphg
    //
    // The hub prefers the API feed when the proxy answers (real titles
    // and dates for the hero list) and falls back to the embedded
    // player when it doesn't. Set the key later and the richer version
    // starts working on its own — nothing here needs changing.
    channelId:         'UC9EwvNr5cfQJW7GRrqCyphg',
    uploadsPlaylistId: 'UU9EwvNr5cfQJW7GRrqCyphg',
  },
  // Marketing asked (deck, 24 Aug) for the stray WordPress mp4 to come
  // off the home page. Turned off here rather than deleted in code, so
  // flipping it back is a one-word change. The YouTube feed fills this
  // box once YOUTUBE_API_KEY is set on the checkfire-jotform Function
  // app — an API key, never the Google account password.
  includeWordPress:  false,
  includeSharePoint: true,
  mediaPortalSite: 'https://checkfireltd.sharepoint.com/sites/CheckFireMediaPortal',
  max: 6,            // how many to show on the home page
  maxAgeMonths: 3,   // only show videos published within the last N months (0 = no limit)
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
// The home-page calendar marks four kinds of date: showroom visits
// (red), product launches (blue), campaign runs (amber) and training
// (green). Users can "Subscribe" to it. On a static site the Subscribe
// button downloads an .ics containing every marked date, which imports
// into Outlook/Google. If you later publish a live SharePoint/Outlook
// calendar, paste its webcal:// or https .ics feed URL below and
// Subscribe will use that instead (a true auto-updating subscription).
HUB_CONFIG.calendar = {
  feedUrl: '',   // e.g. 'webcal://outlook.office365.com/owa/calendar/.../reachcalendar.ics'
};

// ── Jotform — Showroom Booking form ───────────────────────────
// Public booking form embedded as an iframe. NO API key is used or
// stored anywhere in this app — see the security note in jotform.js.
HUB_CONFIG.jotform = {
  formId: '240422414566047',

  // Azure Function proxy that holds the Jotform API key server-side.
  // With this set, the "Upcoming showroom visits" box lists live
  // submissions straight from Jotform as well as anything in the
  // SharePoint list. The key NEVER ships to the browser.
  // Leave '' to read from the SharePoint list only.
  proxyUrl: 'https://checkfire-jotform-fhagcybsfvg5fth8.uksouth-01.azurewebsites.net/api/bookings',
};

// ── Polls (home page) ─────────────────────────────────────────
// A one-question poll card sits under Quick Links on the home page.
// Staff answer it IN THE HUB — the answer is written straight to
// SharePoint, so nobody has to leave for a survey tool.
//
// LIST "Polls" on MarketingHub — the columns marketing built:
//   Title       short code, e.g. "NPD" / "CO2"
//   Question    the question people actually answer
//   Status      choice — only "Open" shows. Draft/Closed stay hidden
//   OpensDate   optional — the card stays hidden until this day
//   ClosesDate  optional — the card hides itself after this day
//   PollURL     optional — adds an "Open the full form →" button
//   Options     OPTIONAL, multi-line, ONE CHOICE PER LINE
//
// The card picks its mode from Options:
//   * Options filled in (2+ lines) → tap-to-vote, results as % bars
//   * Options empty                → free-text answer box
// If more than one poll is Open, the newest one shows.
//
// LIST "Poll Votes" — written by the hub, one row per person per poll:
//   Title    the answer (trimmed to 255 characters)
//   PollId   the Polls item id
//   Voter    the answerer's email
//   Answer   optional multi-line — the full untruncated text. The hub
//            includes it if the column exists and silently drops it if
//            it doesn't, so the list works either way.
//
// WRITE PERMISSION: answering needs the delegated Microsoft Graph scope
// Sites.ReadWrite.All. The hub does NOT ask for it up front — it stays
// read-only until someone actually answers, then requests the extra
// scope once. Without it the card still shows the question and simply
// cannot save. See auth.js and POLLS-SETUP.md.
HUB_CONFIG.polls = {
  list:      'Polls',
  votesList: 'Poll Votes',
};

// ── Team wall (home page) ─────────────────────────────────────
// The internal comms wall is a Viva Engage community embedded in the
// hub. Viva Engage is already part of the CheckFire Microsoft 365
// licence and gives posting, liking, commenting, @mentions and
// notifications for free — none of which a static site can do on its own.
//
// To get the embed URL (note: this is NOT the community's normal
// browser address — it has to be generated):
//   1. Go to https://engage.cloud.microsoft/embed/widget?domainRedirect=false
//   2. Choose "Community feed" on the left and type the community name.
//   3. Click "Get code", then Copy.
//   4. Paste just the iframe's src="..." value into vivaEngageEmbed below.
//
// Microsoft retired the old web.yammer.com JavaScript embed on
// 1 June 2025 — modern embeds come from engage.cloud.microsoft.
//
// Until that is set, the wall falls back to a read-only feed of the
// SharePoint "Comms" list so the section still shows something.
HUB_CONFIG.social = {
  linkedInPageUrl: 'https://www.linkedin.com/company/checkfire/',

  // 6 Aug 2026: repointed from the "All Company" community to the
  // dedicated "Test Marketing Hub" community. All Company is the Viva
  // Engage default group that EVERY user in the tenant belongs to and
  // cannot leave, so every post to the hub wall notified the whole
  // business. Do not point this back at All Company.
  //
  // Generated 31 Jul 2026 from the Embed widget configuration site.
  // Chrome/footer/banner switched off so the feed sits cleanly under the
  // hub's own "Team wall" heading. Flip any of them back to true in the
  // query string if you'd rather have Viva Engage's own framing.
  vivaEngageEmbed: 'https://engage.cloud.microsoft/embed/groups/eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIyMTA0NzY1ODIxMTY1NTY4In0?header=false&footer=false&hideNetworkName=true&theme=light&includeFeedInformation=false',

  // The community itself — used for the "Open in Viva Engage →" link.
  // Community ID: 2104765821165568  ("Test Marketing Hub")
  vivaEngageUrl: 'https://engage.cloud.microsoft/main/groups/eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiIyMTA0NzY1ODIxMTY1NTY4In0/all',

  commsList: 'Comms',    // read-only fallback feed
  commsMax:  8,
};

// ── Landing page images ───────────────────────────────────────
// "Updated landing pages" on the home page reads WordPress *pages*,
// and WordPress pages almost never carry a featured image — which is
// why those cards were bare. Marketing now drop artwork into a folder
// in the marketing Documents library instead, and the hub matches an
// image to a page BY FILENAME.
//
//   Documents ▸ Images for Landing Pages ▸ fire-extinguishers.jpg
//        matches   checkfire.co.uk/fire-extinguishers
//
// Matching is forgiving — case, spaces, underscores, hyphens and the
// extension are all ignored, and a file whose name merely contains the
// page slug (or the other way round) still counts. So
// "Fire Extinguishers.png", "fire_extinguishers.jpg" and
// "01 fire-extinguishers hero.jpg" all land on the same page.
// Anything with no match keeps the current no-image card, so a missing
// picture never leaves an empty grey box.
HUB_CONFIG.landingImages = {
  folder: 'Images for Landing Pages',
  // Where to look for that folder. 'marketing' = the MarketingHub
  // Documents library (where Jess is putting them).
  site: 'marketing',
};

// ── Product Portal front door ─────────────────────────────────
// The Product Portal SharePoint site is organised by CERTIFICATION
// TYPE — DOCs, Kitemark Certificates, Marine Equipment Directive
// (MED), Marine Equipment Regulations (MER), NTA 8133 — with the
// product buried in each filename ("...Declaration of Conformity-CO2-
// AlloySteel.pdf"). Fine for filing, useless if what you actually have
// is a customer asking "send me the paperwork for the 6kg powder".
//
// So the hub indexes the whole site once, then lets people come at it
// from either direction: by PRODUCT or by DOCUMENT TYPE, with a search
// box over every filename. The raw folder browser is still there
// behind "Browse folders" — nothing is hidden.
//
// productTypes: the first pattern that matches a filename wins, so
// order matters (W3E before the general Water rule, for instance).
// Marketing can add a row here without touching any code.
HUB_CONFIG.productPortal = {
  crawlDepth: 3,          // the site is shallow; 3 covers it comfortably
  maxFiles:   400,

  productTypes: [
    { key:'co2',      label:'CO₂',            match:'co2|carbon dioxide' },
    { key:'w3e',      label:'Water W3E',      match:'w3e' },
    { key:'water',    label:'Water',          match:'water|h2o' },
    { key:'foam',     label:'Foam',           match:'foam|\\bff\\b|afff|f3' },
    { key:'powder',   label:'Powder',         match:'powder|\\babc\\b|\\bbc\\b' },
    { key:'wetchem',  label:'Wet chemical',   match:'wetchem|wet chem|wet-chem' },
    { key:'blanket',  label:'Fire blankets',  match:'blanket' },
    { key:'lfx',      label:'LFX',            match:'lfx' },
    { key:'hose',     label:'Hose reels',     match:'hose|en ?671|1866' },
  ],

  // Friendlier names for the top-level folders, and the order the
  // document-type chips appear in. Anything not listed keeps its own
  // folder name and sorts to the end.
  categories: [
    { folder:'DOCs',                               label:'Declarations of Conformity' },
    { folder:'Kitemark Certificates',              label:'Kitemark certificates' },
    { folder:'Marine Equipment Directive (MED)',   label:'MED' },
    { folder:'Marine Equipment Regulations (MER)', label:'MER' },
    { folder:'NTA 8133',                           label:'NTA 8133' },
  ],

  recentCount: 6,
};

// ── Ember — the CheckFire AI assistant ────────────────────────
// Ember lives in a slide-over panel on every page of the hub.
//
// TWO MODES, and it picks whichever is available:
//
// 1. COPILOT STUDIO (the real Ember). Build the agent in Copilot
//    Studio, point its knowledge at the MarketingHub, Product Portal
//    and Media Portal sites, publish it to a custom website, then
//    paste the embed URL below. Copilot Studio honours each person's
//    own SharePoint permissions, so nobody sees a document they
//    couldn't already open, and there is no API key to leak.
//    Step-by-step in EMBER-COPILOT-STUDIO-SETUP.md.
//
// 2. SEARCH MODE (what runs until then). Ember searches all three
//    SharePoint sites through Graph and answers with the documents
//    themselves, opened in-hub. No AI, no cost, and genuinely the
//    fastest way to find a certificate — so the panel is worth
//    shipping before the agent exists.
HUB_CONFIG.ember = {
  enabled: true,
  name: 'Ember',
  tagline: 'CheckFire’s assistant',

  // Paste the Copilot Studio "custom website" embed URL here to switch
  // Ember from search mode to the real agent. Leave '' for search mode.
  copilotEmbedUrl: '',

  // Sites Ember searches in search mode, in the order results group.
  searchSites: [
    { key:'marketing', label:'Marketing Hub',  url:'https://checkfireltd.sharepoint.com/sites/MarketingHub' },
    { key:'product',   label:'Product Portal', url:'https://checkfireltd.sharepoint.com/sites/CheckFireProductPortal' },
    { key:'media',     label:'Media Portal',   url:'https://checkfireltd.sharepoint.com/sites/CheckFireMediaPortal' },
  ],

  // Shown as tap-to-run examples on the empty panel.
  suggestions: [
    'Kitemark certificate for fire blankets',
    'Declaration of conformity CO2',
    'MED certificate LFX',
    'Brand guidelines',
  ],
};
