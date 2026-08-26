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

  // REMOVED 26 Aug 2026 (fix 4) — campaignAssetBlocks.
  //
  // This was a hard-coded list of six folder names used as a "friendly
  // day one" fallback on the Campaign and Product Launch detail pages.
  // In practice it was neither friendly nor a fallback: the tiles drew
  // instantly, said "Open", and did nothing, because no campaign folder
  // in SharePoint is actually called "Data card" or "PR activities".
  // David: "Campaigns doesn't open anything."
  //
  // The detail pages now read the real sub-folders of
  //   Documents/<Campaigns|Launches>/<item name>/
  // and show a plain sentence when there aren't any. Do not put a
  // placeholder list back — an empty state that tells the truth beats a
  // full one that lies.

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

    // 26 Aug 2026, second round. The embedded uploads playlist worked
    // but CheckFire's web filter blocks the YouTube player inside the
    // page, so all anyone saw was a dead frame. Embedding is now gone
    // entirely — the Latest videos box shows the newest upload as a
    // thumbnail card and LINKS OUT to YouTube, which the filter does
    // allow.
    //
    // Where the newest upload comes from: the new checkfire-ai
    // Function app reads the channel's public RSS feed. No API key, no
    // quota, no Google account — see aiProxyUrl below and
    // checkfire-ai-function/DEPLOY.md. Until that's deployed the box
    // falls back to a channel card, and the old /api/videos endpoint
    // on checkfire-jotform is tried first in case its YOUTUBE_API_KEY
    // ever does get set.
    channelId: 'UC9EwvNr5cfQJW7GRrqCyphg',
  },
  // BOTH OFF, 26 Aug 2026 (round 2 fix). The Latest videos box was
  // showing a Media Portal file, so clicking it dropped you into
  // SharePoint instead of YouTube — David's second point. That box is
  // now YouTube only, and renderHeroVideos() filters to YouTube as
  // well, so this can't come back by accident.
  //
  // Flip either back to true if you ever want the box to include
  // WordPress uploads or Media Portal files again.
  includeWordPress:  false,
  includeSharePoint: false,
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

// ── Library front doors (Resources + Product Portal) ──────────
// Both pages run the same component: read the whole folder tree once,
// then offer search, a tile row and type chips over it, with Download
// and Copy link on every row. The raw SharePoint tree stays available
// behind "Browse folders".
//
// PRODUCT PORTAL. The site files by CERTIFICATION TYPE — DOCs,
// Kitemark Certificates, Marine Equipment Directive (MED), Marine
// Equipment Regulations (MER), NTA 8133 — with the product buried in
// each filename ("...Declaration of Conformity-CO2-AlloySteel.pdf").
// Fine for filing, useless when what you have is a customer asking for
// the paperwork on a 6kg powder. `tags` is what puts the product back:
// the FIRST pattern that matches a filename wins, so order matters
// (W3E sits above the general Water rule for exactly that reason).
// Marketing can add a row here without touching any code.
HUB_CONFIG.libraries = {

  product: {
    title: 'Product portal',
    site: 'product',
    hostId: 'pp-index',
    browserId: 'pp-browser',
    browserGridId: 'pp-documents-grid',
    crumbId: 'pp-crumbs',
    backLabel: 'Back to products',
    searchPlaceholder: 'Search every certificate and declaration…',
    tagsLabel: 'By product',
    catsLabel: 'By document type',
    crawlDepth: 3,
    maxFiles: 400,
    recentCount: 6,
    excludeFolders: [],

    tags: [
      { key:'co2',     label:'CO₂',           match:'co2|carbon dioxide' },
      { key:'w3e',     label:'Water W3E',     match:'w3e' },
      { key:'water',   label:'Water',         match:'water|h2o' },
      { key:'foam',    label:'Foam',          match:'foam|\\bff\\b|afff|f3' },
      { key:'powder',  label:'Powder',        match:'powder|\\babc\\b|\\bbc\\b' },
      { key:'wetchem', label:'Wet chemical',  match:'wetchem|wet chem|wet-chem' },
      { key:'blanket', label:'Fire blankets', match:'blanket' },
      { key:'lfx',     label:'LFX',           match:'lfx' },
      { key:'hose',    label:'Hose reels',    match:'hose|en ?671|1866' },
    ],

    categories: [
      { folder:'DOCs',                               label:'Declarations of Conformity' },
      { folder:'Kitemark Certificates',              label:'Kitemark certificates' },
      { folder:'Marine Equipment Directive (MED)',   label:'MED' },
      { folder:'Marine Equipment Regulations (MER)', label:'MER' },
      { folder:'NTA 8133',                           label:'NTA 8133' },
    ],
  },

  // RESOURCES — the marketing Documents library on MarketingHub.
  // Campaigns, Launches and Events are excluded because each already
  // has its own page in the hub; showing them here again just makes
  // the library look like a filing cabinet. Images for Landing Pages
  // is plumbing, not a resource.
  resources: {
    title: 'Marketing library',
    site: 'marketing',
    hostId: 'res-index',
    browserId: 'res-browser',
    browserGridId: 'sp-documents-grid',
    crumbId: 'docs-crumbs',
    backLabel: 'Back to the library',
    searchPlaceholder: 'Search presentations, artwork, guidelines…',
    tagsLabel: 'By kind',
    catsLabel: 'By folder',
    crawlDepth: 3,
    maxFiles: 400,
    recentCount: 6,
    excludeFolders: ['Campaigns', 'Launches', 'Events', 'Images for Landing Pages'],

    // Grouped by what the file IS, since a marketing library is mixed
    // media rather than one product line.
    tags: [
      // Brand sits FIRST on purpose: "CheckFire Brand Guidelines.pdf"
      // is more usefully filed under Brand than under Documents, and
      // the first pattern that matches wins.
      { key:'brand',  label:'Brand & guidelines', match:'brand|guideline|styleguide|style guide|logo|toolkit' },
      { key:'deck',   label:'Presentations', match:'\\.pptx?$|deck|presentation' },
      { key:'doc',    label:'Documents',     match:'\\.(docx?|pdf|rtf)$' },
      { key:'sheet',  label:'Spreadsheets',  match:'\\.(xlsx?|xlsm|csv)$' },
      { key:'image',  label:'Artwork',       match:'\\.(png|jpe?g|gif|webp|svg|ai|eps|psd|indd)$' },
      { key:'video',  label:'Video',         match:'\\.(mp4|mov|webm|m4v)$' },
    ],

    // No fixed list — folder names become the categories as marketing
    // create them, which is what Jess asked for.
    categories: [],
  },

};

// ── Ember — the CheckFire AI assistant ────────────────────────
// Ember lives in a slide-over panel on every page of the hub.
//
// THREE MODES, and it uses the best one available:
//
// 1. CLAUDE (aiProxyUrl set) — a real conversation. The RETRIEVAL
//    happens in the browser, under the signed-in user's own Microsoft
//    token: Ember searches the three SharePoint sites, reads the most
//    relevant documents, and sends only those extracts plus the
//    question to the proxy. The proxy holds the Anthropic key and
//    talks to Claude. Nothing is ever indexed on a server, and Ember
//    can only ever see documents the person asking could already open.
//
// 2. COPILOT STUDIO (copilotEmbedUrl set) — Microsoft's agent, embedded.
//    See EMBER-COPILOT-STUDIO-SETUP.md.
//
// 3. SEARCH (neither set) — no AI: Ember searches and hands back the
//    documents themselves. Costs nothing and still beats hunting
//    through SharePoint.
HUB_CONFIG.ember = {
  enabled: true,
  name: 'Ember',
  tagline: 'CheckFire\u2019s assistant',

  // The checkfire-ai Function app. Set this and Ember becomes a real
  // conversation. Deploy guide: checkfire-ai-function/DEPLOY.md.
  // Example: 'https://checkfire-ai-xxxx.uksouth-01.azurewebsites.net/api'
  aiProxyUrl: 'https://checkfire-ai-eee3fpemdpb0g7g8.uksouth-01.azurewebsites.net/',

  // Alternative brain — a Copilot Studio agent published to a custom
  // website. If BOTH are set, Claude wins; clear aiProxyUrl to switch.
  copilotEmbedUrl: '',

  // How much document text Ember reads before answering. Higher = better
  // answers and a slightly bigger bill; 8 documents at 6k characters is
  // a sensible balance for certificates and datasheets.
  maxDocs: 8,
  maxCharsPerDoc: 6000,

  // Sites Ember searches, in the order results group.
  searchSites: [
    { key:'marketing', label:'Marketing Hub',  url:'https://checkfireltd.sharepoint.com/sites/MarketingHub' },
    { key:'product',   label:'Product Portal', url:'https://checkfireltd.sharepoint.com/sites/CheckFireProductPortal' },
    { key:'media',     label:'Media Portal',   url:'https://checkfireltd.sharepoint.com/sites/CheckFireMediaPortal' },
  ],

  // Shown as tap-to-run examples on the empty panel.
  suggestions: [
    'Which extinguisher for a commercial kitchen?',
    'Do we have a Kitemark certificate for fire blankets?',
    'What changed in the latest powder declaration?',
    'Where are the brand guidelines?',
  ],
};

// ── Training sign-up ──────────────────────────────────────────
// Marketing put a session in the "Training Events" list; staff click
// once in the hub to book onto it. The booking is written to a second
// list so marketing can see who is coming, and the same click drops
// the session into the person's own Outlook calendar.
//
// LIST "Training Signups" on MarketingHub — columns:
//   Title        (single line)  — the session title, copied across
//   SessionId    (single line)  — the Training Events item id
//   Attendee     (single line)  — the person's email
//   AttendeeName (single line, optional)
//   SessionDate  (date, optional)
//
// Writing needs the delegated Microsoft Graph scope
// Sites.ReadWrite.All. The hub stays read-only until someone actually
// books, then asks for the extra scope once — the same pattern the
// poll uses. Without it the button explains rather than failing.
HUB_CONFIG.trainingSignup = {
  list: 'Training Signups',
  enabled: true,
};
