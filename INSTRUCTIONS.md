# CheckFire Marketing Hub — Marketing Feedback Batch (15 Jul 2026)

This batch actions the marketing team's feedback (the "Marketing Hub" annotated PDF)
plus David's three additional asks: add the Product Portal, make everything open
**in-hub**, and add the Twitter-style feed + LinkedIn.

## Upload to GitHub

Repo: `davidsweeneyowen/Marketing-Shareport-Hub` (web UI ▸ Add file ▸ Upload files ▸ Commit).
Upload these **5 files** (all others already live on GitHub are unchanged):

- `index.html`
- `config.js`
- `graph.js`
- `app.js`
- `jotform.js`

No changes to `ui.js`, `auth.js`, `styles.css`, `msal-browser.min.js`, or the Azure
function. Cache-busting versions bumped: config v4, graph v5, app v10, jotform v9.

After committing, hard-refresh the live site (Ctrl+F5). GitHub Pages can take ~1 min.

---

## What changed

### Navigation
- "Products & Launches" renamed to **Product Launches**.
- The red **Book Showroom** button is removed from the top nav. (Booking still opens
  from the calendar's "Book a slot" button and the Trade & Events page.)

### Home page — rebuilt to the mock
- **Latest Videos** no longer sits alone at the top. It now appears as a compact box
  in the hero (right column) and as a full grid lower down the page.
- **Hero right column split into two equal boxes**: *Upcoming Launch* (next/most-recent
  product launch) and *Latest Videos* (newest few, click to jump to the grid).
- **Three equal boxes** below the hero, all the same size: *Upcoming Showroom Visits*,
  *Marketing Calendar*, *Quick Links*.
- **Quick Links** trimmed to the four you asked for: Media Portal, Product Portal,
  Website, CF LinkedIn.
- **Marketing Calendar** now marks three things — showroom visits (red), product
  launches (blue) and campaign runs (amber), with a matching legend. A **Subscribe**
  button sits next to "Book a slot" (see "Calendar subscribe" below).
- **"More from CheckFire"** is split into two scrollable carousels with ‹ › arrows:
  *Latest Blogs* (WordPress posts) and *Updated Landing Pages* (WordPress pages,
  newest-updated first).
- The **Twitter-style in-house comms feed** and **LinkedIn** panel sit near the bottom
  (both hide themselves until they have content — see setup below).

### Product Portal (second SharePoint site) — everything opens in-hub
- New site added: `https://checkfireltd.sharepoint.com/sites/CheckFireProductPortal`.
- Resources page now has three tabs: **Marketing Library**, **Product Portal**,
  **Useful links**.
- Both library tabs are a proper **in-hub file browser**: click a folder to drill in
  (with a breadcrumb trail), click a file to preview it in the pop-up viewer. Users
  never bounce out to SharePoint. This reuses the same delegated permissions the hub
  already has (Sites.Read.All / Files.Read.All) — **no new admin consent needed**, as
  long as the person is a member of the Product Portal site.

### Campaigns page — rebuilt to the mock
- Clicking a campaign card opens a **detail view**: black hero (title, status pill,
  dates, "View landing page" link), a **metrics bar** (Emails sent / Social media posts
  / Blogs / PR activity), and **asset blocks** (Infographic, Email signature, Email,
  Data card, Social media assets, PR activities).
- Clicking an asset block opens the matching document(s) **in-hub** (single file opens
  straight into the viewer; multiple files list underneath the blocks).
- The same detail layout is wired to the **Product Launches** page, ready for Josh/Lowri
  to decide what goes on it.

---

## What you (marketing) need to set up in SharePoint

These features are built and safe — each one shows a friendly empty state until the
data exists, so nothing breaks if it's not set up yet.

1. **Campaign asset blocks** — for each block to open a document, the Documents library
   needs this folder structure:
   `Documents / Campaigns / <Campaign name> / <Block name>`
   e.g. `Documents/Campaigns/Hidden Risk of Lithium-Ion/Infographic/…`.
   The `<Campaign name>` folder is matched to the campaign's Title (case-insensitive).
   If you'd rather name the folder differently, add a **CampaignFolder** column to the
   Campaigns list and put the folder name there. Block folder names are set in
   `config.js` (`campaignAssetBlocks`) — rename/add freely.

2. **Campaign metrics** (your "how do we manage this?" question) — the four numbers read
   from columns on the **Campaigns** list. Add number columns named **EmailsSent**,
   **SocialPosts**, **Blogs**, **PRActivity**. Then marketing just edits the list item to
   update the numbers — no code needed. Until the columns exist they show 0.

3. **Twitter-style in-house comms feed** — create a SharePoint list called **Comms** on
   the MarketingHub site with columns: **Message** (multi-line text), **Author** (or use
   Title), **Handle**/Team, **Date**, and optionally **Link**. Posts appear newest-first.
   Panel stays hidden while the list is empty.

4. **LinkedIn** — there's no free auto-updating LinkedIn page feed, so the panel shows
   specific posts you choose. In LinkedIn: open a post ▸ ••• ▸ *Embed this post* ▸ copy
   the `src="…"` URL ▸ paste it into `HUB_CONFIG.social.linkedInEmbeds` in `config.js`.
   The CF LinkedIn quick link and "View page" link already point to your company page.

5. **Calendar subscribe** — the Subscribe button currently downloads an `.ics` file of
   all marked dates, which imports into Outlook/Google/Apple. For a *live* auto-updating
   subscription, publish a SharePoint/Outlook calendar and paste its `webcal://` or
   `.ics` feed URL into `HUB_CONFIG.calendar.feedUrl` — Subscribe will then use that.

---

## Open items / advice

- **Campaign metrics management**: recommended approach is the Campaigns-list columns
  above (option 2) — keeps it fully in marketing's hands with no developer involvement.
- **Product Launches detail page**: the scaffold is live and identical to Campaigns.
  Once Josh/Lowri decide the content, we point its asset blocks at a
  `Documents/Launches/<Launch name>/…` structure (already wired).
- **Product Portal library**: assumed the site's default **Documents** library. If the
  portal uses a differently-named library, tell me and I'll set it in `config.js`.

## Assumptions made (flag anything you'd prefer different)

- Hero "split into two boxes" read as *Upcoming Launch* + *Latest Videos*; the three
  boxes below became *Visits / Calendar / Quick Links* (all equal), which matches the
  five annotations on that slide.
- "Latest blogs" = WordPress posts; "Updated landing pages" = WordPress pages. If
  landing pages live somewhere else, we can repoint them.
