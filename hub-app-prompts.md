# CheckFire Marketing Hub — Copy-Paste Prompts for AI Models

Use these in order — each is self-contained and includes the context the model needs. Paste the relevant files alongside each prompt. Ordered by priority.

---

## Prompt 1 — Rebuild the missing UI core (DO THIS FIRST)

> I'm building a vanilla JavaScript single-page intranet hub (no frameworks, no build step — plain JS files loaded via script tags at the bottom of body). My index.html calls these functions but the file that defined them was lost. Recreate a single file called `ui.js` containing exactly these, matching the call signatures used in my HTML:
>
> 1. `showPage(id, idx)` — pages are `<main>` elements with ids `page-home`, `page-launches`, `page-campaigns`, `page-trade`, `page-training` and class `page`. The active page has class `active` (CSS handles display/animation). It's called like `showPage('launches', 1)`. It should also scroll to top and call `updateNavActive(id)` if that function exists.
> 2. `showToast(msg)` — my CSS already has `.toast` (fixed, bottom-centre) and `.toast.show`. Create the element once, reuse it, show for ~2.5 seconds.
> 3. `switchTab(btn, name)` — for the launches page: tabs are `.ltab` buttons inside a `.ltabs` nav; panes have ids `tab-{name}` with class `ltab-pane`, active pane gets `.active`. Toggle `active` on the clicked button's siblings too.
> 4. `switchTrainingTab(btn, name)` — same pattern but panes are `ttab-{name}` and only scoped within the training page's own `.ltabs`.
> 5. `loadResourcesData()` — stub for now: if `window.HUB_DEMO_MODE` is true or no auth token exists, render a friendly "Connect SharePoint to see live documents" message into `#sp-documents-grid`; otherwise leave a TODO comment where the Graph call will go.
>
> Constraints: defensive null checks on every getElementById (some elements only exist on certain pages), no globals beyond the functions themselves, UK English in any user-facing strings. Output the complete file only.

---

## Prompt 2 — Build the real SharePoint/Graph integration

> I have a vanilla JS intranet hub that authenticates users via OAuth PKCE against Entra ID with scopes `Sites.Read.All Files.Read.All User.Read`. An async function `getAccessToken()` already exists and returns a bearer token (or null). My config object is:
>
> ```js
> const HUB_CONFIG = {
>   sharepointSite: 'https://checkfireltd.sharepoint.com/sites/CheckFireMediaPortal',
>   lists: { launches: 'Product Launches', campaigns: 'Campaigns', events: 'Events' },
>   documentsLibrary: 'Documents',
> };
> ```
>
> List columns: **Product Launches** = Title, LaunchDate, SKU, Status, Description, RRP, LinkURL. **Campaigns** = Title, CampaignType, Status, StartDate, EndDate, Budget, Channels, Region, LinkURL. **Events** = Title, EventDate, Location, EventType, Status, LinkURL.
>
> Write a complete `graph.js` (vanilla JS, no imports) that:
> 1. Resolves the site ID once via `GET https://graph.microsoft.com/v1.0/sites/{hostname}:{server-relative-path}` and caches it.
> 2. `fetchListItems(listName)` — gets items with `?expand=fields`, returns the `fields` objects. Handle 404 (list not found) with a thrown error naming the list.
> 3. `fetchLibraryFiles()` — lists the document library's root children via the drive API, returning name, size, lastModifiedDateTime, webUrl, and file type.
> 4. Renderers: `renderLaunches(items)` into `#sp-launches-list` (rows: title, SKU, formatted LaunchDate as "2 Jun 2026", status badge, link), `renderEvents(items)` into `#sp-events-list`, `renderDocuments(files)` into `#sp-documents-grid` (file icon by extension, human-readable size, "Open" link).
> 5. A `loadSharePointData()` orchestrator: skips entirely in demo mode (`window.HUB_DEMO_MODE`), runs the three fetches in parallel with `Promise.allSettled`, renders failures as inline error messages (e.g. "Couldn't load launches: List not found — check config.js") rather than breaking the page.
> 6. SessionStorage caching with a 5-minute TTL, keyed per list.
> 7. Escape ALL field values with an `escHtml()` helper before inserting into innerHTML, including values used in attributes.
>
> Dates display in en-GB format. UK English. Output the complete file.

---

## Prompt 3 — Migrate auth to MSAL.js with true silent SSO

> My vanilla JS SPA currently uses a hand-rolled OAuth2 PKCE flow (I'll paste auth.js). I have msal-browser.min.js available locally. Rewrite auth.js to use MSAL Browser instead, keeping the same public surface so the rest of the app doesn't change: `initAuth()` (returns true if signed in), `getAccessToken()` (returns a Graph token or null), `signIn()`, `signOut()`, and it must still call `showSignedIn(user)`, `showSignInPage()`, `showDemoMode()` exactly as the current file does, and set `window.AUTH = { token, account }` for other scripts.
>
> Requirements:
> - Config: tenantId and clientId come from `HUB_CONFIG`; redirect URI = `window.location.origin + window.location.pathname`.
> - On init: try `ssoSilent` first so staff on work devices with an existing Microsoft session NEVER see a login screen; fall back to showing the sign-in page only if silent fails with InteractionRequired.
> - Use redirect flow (not popup) for the interactive fallback.
> - `getAccessToken()` uses `acquireTokenSilent` with scopes `User.Read Sites.Read.All Files.Read.All`, falling back to redirect on InteractionRequired.
> - Demo mode: if `HUB_CONFIG.clientId === 'YOUR_CLIENT_ID'`, call `showDemoMode()` and skip MSAL entirely.
>
> Explain in two sentences at the top (as a comment) why ssoSilent needs the login_hint or a domain_hint to be reliable, and implement whichever you recommend. Output the complete file.

---

## Prompt 4 — Azure Function proxy for the Jotform API key

> Security fix: my static web app currently calls the Jotform EU API directly from the browser with a full account API key in client-side config — anyone can view source and read all booking submissions. I'm hosting on Azure Static Web Apps, which supports managed Azure Functions in an `/api` folder.
>
> Write me:
> 1. `api/showroom-bookings/index.js` — a Node.js Azure Function (v4 programming model) that calls `https://eu-api.jotform.com/form/{FORM_ID}/submissions?apiKey={KEY}&limit=100&orderby=created_at` with the key and form ID read from environment variables (`JOTFORM_API_KEY`, `JOTFORM_FORM_ID`). It must strip the response down to ONLY the fields the frontend needs per submission: id, booking date, company name, account manager email, number of customers, arrival time — extracted from Jotform's `answers` object (the appointment answer is `{ date: "2026-06-10 03:30" }` or similar; field names are `appointment`, `companyName`, `pleaseConfirm`, `customerNames12`, `numberOf`, `arrivalTime`). Never forward raw submissions. Cache the result in memory for 60 seconds.
> 2. The `staticwebapp.config.json` route addition so `/api/*` requires the `authenticated` role.
> 3. The three-line change to my frontend `fetchShowroomSubmissions()` to call `/api/showroom-bookings` instead.
> 4. The Azure CLI command to set the two app settings.
>
> Keep the function dependency-free (use global fetch, Node 18+).

---

## Prompt 5 — Security & XSS hardening pass

> Review the attached vanilla JS files (app.js, jotform.js, graph.js) for injection vulnerabilities. Known issues to fix and verify: (1) WordPress `post.link` and `post.image` are interpolated into `href` and `style="background-image:url(...)"` attributes without escaping or URL validation — validate they're https URLs on expected domains before use; (2) Jotform answer values (`arrivalTime`, `numCustomers`, `amEmail`) render into innerHTML unescaped in the visits list; (3) check every other `innerHTML` assignment and template literal for unescaped dynamic values, including attribute contexts where `escHtml` alone is insufficient. Don't refactor working logic — produce a minimal diff-style list of changes (file, function, before/after) plus a corrected `escAttr()`/`safeUrl()` helper pair I can drop in.

---

## Prompt 6 — Hash-based routing and working search

> My vanilla JS SPA switches pages via `showPage(id, idx)` toggling `.page.active` on `<main>` elements (ids: page-home, page-launches, page-campaigns, page-trade, page-training). Two improvements, no frameworks:
>
> 1. **Hash routing:** map `#/home`, `#/launches`, `#/campaigns`, `#/trade`, `#/resources` to pages. On load, open the page in the hash (default home). `showPage` updates the hash; `hashchange` (back/forward buttons) updates the page. Avoid infinite loops between the two. Keep the existing `showPage(id, idx)` signature working for all existing onclick handlers.
> 2. **Search:** wire up `#hub-search-input` to a lightweight client-side search. On page load, build an index from each page's headings and card titles (querySelector over `.page h1, h2, h3, .camp-name, .show-name, .module-name, .link-name, .report-name`). Typing 2+ characters shows a dropdown of up to 8 results under the search box (create the dropdown element; basic styling inline or as a small CSS block I can paste into styles.css). Clicking a result navigates to that page via the router and briefly highlights the matched element. Escape key and outside-click close it.
>
> Output one file `router-search.js` plus the small CSS block. Defensive null checks throughout.

---

## Prompt 7 — Accessibility pass

> Audit the attached index.html for keyboard and screen-reader accessibility and give me concrete patches. Known issues: nav links are `<a>` elements with onclick but no href (not focusable); tab buttons (`.ltab`) lack `role="tab"`, `aria-selected` and arrow-key support; the booking modal (`#booking-modal`) has no focus trap, no `role="dialog"`/`aria-modal`, and focus isn't returned to the trigger on close; clickable cards are `<div>`/`<article>` with onclick only; icon-only buttons lack labels. For each: show the exact before/after HTML and any small JS needed. Target WCAG 2.1 AA. Don't redesign anything visually — attribute and behaviour changes only.

---

## Prompt 8 — Reusable review prompt (for any future change)

> You are reviewing a change to the CheckFire Marketing Hub: a vanilla-JS, no-build-step intranet SPA hosted on Azure Static Web Apps. Hard rules for this codebase: no frameworks or npm dependencies in the frontend; all secrets live in Azure Function app settings, never in client files; every dynamic value inserted via innerHTML must pass through escHtml/escAttr and URLs through safeUrl; all user-facing text in UK English; dates formatted en-GB ("2 Jun 2026"); currency in £; data that changes more than quarterly must come from SharePoint lists, not hardcoded HTML; every getElementById result null-checked. Review the attached diff against these rules plus general correctness, and flag anything that breaks the existing public function contracts (initAuth, getAccessToken, showPage, showToast, loadShowroomData).

---

## Tips for using these

Paste the actual files with each prompt — the prompts describe the contracts, but the model does better seeing real code. Run Prompt 1 first and test navigation works before anything else; most of the app's buttons depend on it. Prompts 2 and 4 are the two that turn this from a prototype into the real product. Prompt 8 is for the long haul — paste it at the top of any future "change this" conversation so every model stays inside the same guardrails.
