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

// ── Latest updates — the corner dock ──────────────────────────
// 1 Sep 2026 — marketing: "How can we manage this? Ideally, I would
// like it to include only the most urgent updates."
//
// It used to show every launch and every campaign, newest eight, which
// is a feed rather than something anyone controls. It now shows only
// what marketing have TICKED.
//
// WHAT TO ADD IN SHAREPOINT (2 minutes, once):
//   Product Launches ▸ + Add column ▸ Yes/No ▸ name it exactly  Pinned
//   Campaigns        ▸ + Add column ▸ Yes/No ▸ name it exactly  Pinned
// Tick Pinned on the one or two things people need to know about. Untick
// it and the update disappears — nothing to delete.
//
// With nothing ticked the dock hides completely rather than falling back
// to the old firehose: an empty corner is the honest answer to "there's
// nothing urgent". Set `requirePinned:false` if you'd rather it drop back
// to the newest few while the column is being added.
//
// Each update also carries a link straight to that launch's or
// campaign's asset folder — marketing: "when we're announcing a new
// launch, to include a link to the folder containing all of the assets".
HUB_CONFIG.updates = {
  column: 'Pinned',        // Yes/No column on Product Launches + Campaigns
  requirePinned: true,     // false = show the newest `max` when none ticked
  max: 5,
  linkToAssets: true,
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
//
// 1 Sep 2026 — marketing: "Can this page please go as it was. Divided
// into: Exhibitions / Customer events / Training. Don't want that to
// feel it's only for FSE." Documents ▸ Events held nothing but FSE
// 2025/2026/2027, so the page had no way to look like anything else.
//
// The split comes from SharePoint, not from code: create a folder for
// each category under Documents ▸ Events and put the event folders
// inside it —
//
//   Documents ▸ Events ▸ Exhibitions ▸ FSE 2026
//   Documents ▸ Events ▸ Customer Events ▸ Open Day 2026
//
// Any event folder still sitting loose at the top level is shown under
// `fallback` below, so nothing disappears while the folders are being
// moved. `aliases` means the category folder can be named the way it
// reads best — "Exhibitions & Shows" still matches.
HUB_CONFIG.tradeEvents = {
  folder: 'Events',
  categories: [
    { key:'exhibitions', folder:'Exhibitions',
      aliases:['Exhibitions & Shows','Trade Shows','Shows','Exhibition'],
      label:'Exhibitions',
      sub:'The shows we exhibit at — stand plans, artwork, forms and the packs that go with them.' },
    { key:'customer', folder:'Customer Events',
      aliases:['Customer Event','Customer','Open Days','Open Day'],
      label:'Customer events',
      sub:'Open days, demonstrations and everything we run for customers.' },
  ],
  fallback: 'exhibitions',
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
//
// 1 Sep 2026 — "Image is still not pulling through" on the Fire
// Equipment Suppliers card. The artwork WAS there; the matching was too
// literal. The folder is "Fire Equipment Supplier Landing Page" and the
// page is /fire-equipment-suppliers, so neither name contained the
// other and the old substring rule found nothing. Matching is now done
// on WORDS: the words below are stripped as noise first, a trailing "s"
// is ignored, and the picture sharing the most words with the page
// wins. "Fire Equipment Supplier Landing Page" and "Fire Equipment
// Suppliers" now share three words and match.
HUB_CONFIG.landingImages = {
  folder: 'Images for Landing Pages',
  // Where to look for that folder. 'marketing' = the MarketingHub
  // Documents library (where Jess is putting them).
  site: 'marketing',
  // Words that say nothing about WHICH page this is.
  noiseWords: ['landing','page','pages','image','images','hero','banner',
               'main','cover','final','new','copy','v1','v2','checkfire','cf'],
  // How many real words must line up before it counts as a match.
  minWordMatch: 1,
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
    searchPlaceholder: 'Search every datasheet, certificate, MSDS and notification…',
    tagsLabel: 'By product',
    catsLabel: 'By document type',
    crawlDepth: 3,
    maxFiles: 1400,
    recentCount: 6,
    excludeFolders: [],

    // ── THE HOME FOR EVERYTHING ─────────────────────────────
    //
    // 1 Sep 2026, second round. David: "I have done the product portal,
    // you should have more than enough data in the SharePoint to turn
    // this into the home for everything."
    //
    // He was right, and the hub was looking in one place. Two things
    // were wrong:
    //
    // 1. The six Product Change Notifications on the Product Portal
    //    site are not in its Documents library at all — they are in
    //    **FormServerTemplates**, a system library, because they were
    //    dropped onto a page rather than into the library. A normal
    //    library crawl cannot see them, which is why the portal looked
    //    like it had none. `allLibraries` fixes that: every library on
    //    the site gets read, system ones included.
    //
    // 2. The current datasheets, the PIF/MSDS toolkits, the launch
    //    packs and the product training all live on OTHER sites. They
    //    are now sources here, so the Product Portal page is the one
    //    place to look.
    //
    // Nothing is copied or moved. Each source is read live, under the
    // signed-in person's own permissions — if they can't open it in
    // SharePoint they won't see it here either.
    //
    // sales.marketing is a big, mixed site with commercial folders on
    // it (pipeline, budgets, tenders), so it is deliberately read by
    // NAMED ROOTS only, never from the top. Add a root here to add it
    // to the portal; there is no wildcard on purpose.
    sources: [
      {
        key: 'portal', label: 'Product Portal',
        site: 'https://checkfireltd.sharepoint.com/sites/CheckFireProductPortal',
        allLibraries: true, depth: 4, max: 400,
      },
      {
        key: 'media', label: 'Media Portal',
        site: 'https://checkfireltd.sharepoint.com/sites/CheckFireMediaPortal',
        library: 'Documents', depth: 5, max: 500,
        roots: [
          '01. Fire Extinguishers',
          '06. Bridgehill Fire Blankets',
          '07. Alarms',
          '08. Hose Reels',
        ],
      },
      {
        key: 'sales', label: 'Sales & Marketing',
        site: 'https://checkfireltd.sharepoint.com/sites/sales.marketing',
        library: 'Documents', depth: 3, max: 500,
        roots: [
          '04. Product Change Notifications',
          'Product Documents',
          'Service Manuals',
          '01. Marketing/08. PDF PIF, Data Sheets, MSDS Sheets & Toolkits',
          '01. Marketing/09. Product Launch Packs',
          '01. Marketing/07. NPD',
          '01. Marketing/14. Product Training',
        ],
      },
    ],

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

    // With five sources feeding one page, "the top-level folder" is no
    // longer a useful grouping — the same kind of document sits under a
    // different folder name on every site. So a category can now be
    // matched on the FILE PATH AND NAME with `match`, and the first
    // rule that matches wins. `folder` rules still work and are tried
    // first, so nothing that used to group correctly stops.
    categories: [
      { folder:'DOCs',                               label:'Declarations of Conformity' },
      { folder:'Kitemark Certificates',              label:'Kitemark certificates' },
      { folder:'Marine Equipment Directive (MED)',   label:'MED' },
      { folder:'Marine Equipment Regulations (MER)', label:'MER' },
      { folder:'NTA 8133',                           label:'NTA 8133' },

      { match:'product change notification|\\bpcn\\b',  label:'Product change notifications' },
      { match:'declaration of conformity|\\bdoc\\b',    label:'Declarations of Conformity' },
      { match:'kitemark',                             label:'Kitemark certificates' },
      { match:'\\bmsds\\b|safety data sheet|\\bsds\\b', label:'MSDS & safety data' },
      { match:'\\bpif\\b|product information file', label:'PIF' },
      { match:'data ?sheet',                          label:'Data sheets' },
      { match:'launch pack',                          label:'Launch packs' },
      { match:'service manual|instruction|user guide|manual', label:'Manuals & instructions' },
      { match:'training',                             label:'Product training' },
      { match:'\\bnpd\\b|new product',              label:'New product development' },
      { match:'toolkit',                              label:'Toolkits' },
      { match:'certificat|approval|\\bced\\b|\\ben ?3\\b', label:'Certificates & approvals' },
      { match:'brochure|flyer|leaflet',               label:'Brochures' },
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

    // 1 Sep 2026 — marketing: "This feels very chaotic at the moment.
    // Can we just have a list as before with documents added to the
    // SharePoint." With three files in the library, a tile row, a chip
    // row and a "recently updated" rail were three ways of saying the
    // same three things. `simple` drops all of that: a search box and
    // the documents, grouped by the folder they sit in. Set it to
    // false to get the faceted view back.
    simple: true,
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
  // 1 Sep 2026 — David: "we'd like to change the name to Josh 2.0".
  // Changing it here changes the launcher, the panel, the placeholder
  // and the greeting. The Function app also needs redeploying for the
  // assistant to CALL itself Josh 2.0 in its own answers — the name is
  // sent with every question as `persona`, and a Function that predates
  // this simply ignores it.
  name: 'Josh 2.0',
  tagline: 'CheckFire\u2019s assistant',

  // 1 Sep 2026 — David pointed at the Jotform agent marketing had
  // already built ("Josh 2.0 · Product Specialist", 35 conversations
  // going back to June 2025) and asked for the hub's chat to work like
  // it. The role sits under the name in the panel header, exactly as it
  // does there: it tells you what to ask before you've typed anything.
  role: 'Product Specialist',

  // The checkfire-ai Function app. Set this and Ember becomes a real
  // conversation. Deploy guide: checkfire-ai-function/DEPLOY.md.
  // Example: 'https://checkfire-ai-xxxx.uksouth-01.azurewebsites.net/api'
  aiProxyUrl: 'https://checkfire-ai-eee3fpemdpb0g7g8.uksouth-01.azurewebsites.net/api',

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
  // The opening taps. The Jotform agent offers two ("Explore products",
  // "Learn more") — short enough to read at a glance, broad enough to
  // start anywhere. Follow-ups after that are written by the assistant
  // for the answer it just gave (see the SUGGEST line in ember.js).
  suggestions: [
    'Explore our products',
    'Find a certificate or datasheet',
    'What\u2019s launching next?',
    'Draft something for me',
  ],

  // ── The product catalogue ──────────────────────────────────
  // 2 Sep 2026 — David, after Josh said the only Commander extinguisher
  // item was a STAND: "I thought having the API key would help with
  // things like this? Our Commander range is everything that says
  // Commander, right."
  //
  // Right, and the API key was never going to do it. The key buys the
  // ability to read and reason; it does not buy any knowledge of what
  // CheckFire sells. Josh only ever knows what the hub hands him for a
  // question, and what it was handing him was the Product Launches and
  // Campaigns lists — this quarter's marketing ACTIVITY, not the
  // product range. So he answered about a launch of a stand, because a
  // stand launch is what the lists contain.
  //
  // The range itself has always been in the Media Portal, filed exactly
  // the way anyone would want:
  //
  //   01. Fire Extinguishers ▸ Commander Fire Extinguishers ▸ CO2 /
  //     Foam / Powder / Water / WetChemical / Wheeled Units
  //   02. Commander Cabinets · 03. Stands · 05. Fire Blankets ·
  //   06. Bridgehill · 07. Alarms · 08. Hose Reels · …
  //
  // That folder tree IS the catalogue. It is read once per session, to
  // the depth below, and handed to Josh with every question — so "tell
  // me about our Commander range" is answered from the range, and
  // "what other Contempo is there" has an answer.
  //
  // Names only. No file is opened for this, and it costs one short
  // burst of Graph calls the first time somebody asks.
  catalogue: {
    site: 'https://checkfireltd.sharepoint.com/sites/CheckFireMediaPortal',
    library: 'Documents',
    depth: 3,        // category ▸ range ▸ type
    max: 400,
  },

  // ── Answering like a conversation, not a search box ────────
  // 1 Sep 2026 — David: "we need to ensure it's working like an
  // actually AI chat bot. Like talking to you."
  //
  // It already had Claude behind it. What made it feel like a search
  // tool was that EVERY message went through a full SharePoint search
  // first — five seconds of "looking through SharePoint…" before a
  // reply to "morning" or "make that shorter". It now only searches
  // when the question is actually about a document, and streams the
  // answer a word at a time when the Function app supports it.
  //
  // Set skipSearchWhenChatty:false to go back to always searching.
  skipSearchWhenChatty: true,
  stream: true,

  // Short. The Jotform agent answers a product question in one line —
  // "For electrical fires, use a CO2 extinguisher or a dry powder
  // extinguisher; if you want, I can also help you choose the right one
  // for home, office, or vehicle use." — and that brevity is most of why
  // it reads as a conversation. 700 tokens is roughly a long paragraph:
  // enough for a real answer, not enough to lecture. It lifts by itself
  // when someone asks for something long (see maxTokensLong).
  maxTokens: 700,
  maxTokensLong: 2000,
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

// ── Product Portal front door ─────────────────────────────────
// 1 Sep 2026. The product team asked for the whole Product Portal to
// move into the Marketing Hub, "ideally keeping the organisation the
// same". Their list of what the portal has that the hub's page didn't:
//
//   Product Change Notifications · Links to PIF · Sample Request Sheet
//   New Product Request Sheet · launch countdown and upcoming dates
//   · a feedback form · datasheets/MSDS
//
// The certificate index underneath is untouched — same products, same
// document types, same search. These are bands added above it.
//
// EACH BAND IS A FOLDER. The hub reads the top-level folders of the
// Product Portal library it already crawls, so a band appears the
// moment its folder exists and is simply absent until then — no empty
// shells, no "coming soon". Aneta/Jess create these on
// sites/CheckFireProductPortal ▸ Documents:
//
//   Product Change Notifications
//   PIF
//   Sample Requests
//   New Product Requests
//   Data Sheets and MSDS
//
// `aliases` keeps it forgiving — "PIFs", "Product Information Files"
// and "Links to PIF" all match the PIF band.
HUB_CONFIG.productPortal = {
  // A band is a VIEW over the one index below it, not a second copy of
  // the files. It appears when the index actually contains that kind of
  // document and is absent otherwise — no empty shells.
  //
  // `cats` are category labels from HUB_CONFIG.libraries.product
  // .categories, which are matched on the document's own path and name,
  // so a band works no matter which site the file came from. `folder`
  // and `aliases` are the older folder-name route, still honoured for
  // the two request sheets, which are a folder rather than a kind.
  sections: [
    { key:'pcn',     label:'Product change notifications',
      cats:['Product change notifications'],
      folder:'Product Change Notifications',
      aliases:['Product Change Notification','PCN','Change Notifications'],
      desc:'Every notified change to a product, newest first.' },

    { key:'data',    label:'Datasheets & MSDS',
      cats:['Data sheets','MSDS & safety data'],
      folder:'Data Sheets and MSDS',
      aliases:['Datasheets','Data Sheets','MSDS','SDS'],
      desc:'Technical data sheets and safety data sheets.' },

    { key:'certs',   label:'Certificates & declarations',
      cats:['Declarations of Conformity','Kitemark certificates','MED','MER',
            'NTA 8133','Certificates & approvals'],
      desc:'Conformity, Kitemark, marine and NTA paperwork.' },

    { key:'pif',     label:'Links to PIF',
      cats:['PIF'],
      folder:'PIF',
      aliases:['PIFs','Links to PIF','Product Information Files','Product Information File'],
      desc:'The product information file for each product.' },

    { key:'manuals', label:'Manuals & instructions',
      cats:['Manuals & instructions'],
      desc:'Service manuals, user guides and fitting instructions.' },

    { key:'launch',  label:'Launch packs',
      cats:['Launch packs'],
      desc:'Everything that went out with each product launch.' },

    { key:'training',label:'Product training',
      cats:['Product training'],
      desc:'Training material for the range.' },

    { key:'npd',     label:'New product development',
      cats:['New product development'],
      desc:'What is coming, and what is being worked on.' },

    { key:'samples', label:'Sample request sheet',
      folder:'Sample Requests',
      aliases:['Sample Request Sheet','Sample Request','Samples'],
      desc:'Request a sample for a customer.' },

    { key:'npr',     label:'New product request sheet',
      folder:'New Product Requests',
      aliases:['New Product Request Sheet','New Product Request','NPD Requests'],
      desc:'Put a product forward for the range.' },
  ],

  // Launch countdown and "dates to look out for", from the same
  // Product Launches list the rest of the hub reads.
  upcomingCount: 4,

  // Fast feedback on a product. Paste a Microsoft Form or Jotform URL
  // here and the box appears at the foot of the page; leave it empty
  // and there is no box.
  feedbackUrl: '',
  feedbackTitle: 'Feedback on a product',
  feedbackSub: 'Something wrong with a datasheet, a certificate out of date, or a product you keep being asked for? Tell the product team.',
};
