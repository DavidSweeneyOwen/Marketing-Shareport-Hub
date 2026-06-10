# CheckFire Marketing Hub — Code Review & Improvement Plan

*Reviewed 10 June 2026 · files: index.html, styles.css, app.js, auth.js, config.js, graph.js, jotform.js, staticwebapp.config.json, DEPLOY.md*

---

## What's genuinely good

The design system is strong — Fraunces/Manrope pairing, skeleton loaders everywhere, consistent card language, and the CSS is well organised with sensible breakpoints. The hand-rolled PKCE flow in auth.js is done properly (S256 challenge, state validation, URL cleanup). The Jotform calendar logic is solid: month navigation, free-day counting, the mini calendar on the Trade page, and the four-state handling for the visits panel (loading / error / empty / loaded) shows real care. DEPLOY.md is one of the better deployment guides I've seen for an internal tool — a non-developer could follow it. Demo mode as a fallback when config isn't set is a nice touch.

---

## Critical issues (the app is currently broken)

### 1. A whole JavaScript file appears to be missing
`showPage()`, `showToast()`, `switchTab()`, `switchTrainingTab()`, `loadResourcesData()` and `setupSignInButton()` are called throughout index.html but **defined nowhere** in the uploaded files. The CSS has matching `.page.active` and `.toast` rules, so this code existed at some point — it looks like a `main.js` (or similar) was lost or never committed.

**Effect right now:** clicking any nav link, tab, or toast-button throws a ReferenceError. Navigation between pages doesn't work at all. This is the single highest-priority fix.

### 2. graph.js contains no Microsoft Graph code
Despite the name, the scopes requested at sign-in (`Sites.Read.All`, `Files.Read.All`), the list names in config.js, and the column schema in DEPLOY.md — there isn't a single Graph API call in the codebase. The skeletons in `#sp-launches-list` and `#sp-documents-grid` will spin forever. The SharePoint integration is the heart of the product and it doesn't exist yet.

### 3. Jotform API key will be exposed client-side
config.js ships to every browser. Once you paste your Jotform API key in, anyone who views source has **full API access to your Jotform account** — including all booking submissions with customer names, company names and emails. And with the redirect URI pointing at public GitHub Pages, that's the open internet, not just staff. This needs a server-side proxy (an Azure Function or SWA managed API) before go-live. Non-negotiable.

### 4. Hosting identity crisis
Three different stories in the code: config.js redirects to **GitHub Pages** (`davidsweeneyowen.github.io`), DEPLOY.md describes **Azure Static Web Apps**, and staticwebapp.config.json declares `allowedRoles: ["authenticated"]` — which relies on SWA's *built-in* auth that's never configured. Deployed to SWA as-is, that config would 401 every request and redirect to a `/login` route that doesn't exist. On GitHub Pages it's silently ignored. Pick one host and align all three files.

### 5. The account-manager pre-fill never fires
jotform.js checks `window.AUTH.account.mail` — but `AUTH` is declared with `const` in auth.js, so it never becomes a property of `window`. `window.AUTH` is always `undefined` and the booking form never pre-fills. One-line fix: `window.AUTH = AUTH;` in auth.js.

### 6. MSAL is shipped but never used
index.html has an "MSAL.js" comment in the head with no script tag, and msal-browser.min.js sits unused in the folder. Meanwhile auth.js claims "silent SSO first" in its header comment but never attempts a `prompt=none` flow — users always see the sign-in screen. Either adopt MSAL properly (recommended — it handles silent SSO, token caching and renewal for free) or delete the dead file.

---

## Smaller but worth fixing

**Security/XSS:** `post.link` and `post.image` from WordPress are interpolated into `href` and `style` attributes unescaped (titles/excerpts are escaped — good). `arrivalTime` and `numCustomers` from Jotform render unescaped in the visit meta line. Refresh token in sessionStorage is acceptable for an internal tool but MSAL would manage this better.

**Duplicate fetch:** `fetchWordPressNews()` is called twice on every load (hero grid + news section). Cache the promise.

**Dead/duplicate sign-in code:** the inline `signIn()` in index.html calls the non-existent `setupSignInButton()`. It happens to be overwritten when auth.js loads later, so sign-in works — but it's a landmine. Delete the inline version.

**Hardcoded content masquerading as live:** the launches hero (ProGuard), price facts, launch team, readiness checklist, all six campaign cards, all four events, training modules, pathways, reports, surveys, the countdown date, the homepage product news, and the "/ 22" free-days denominator. Fine for a prototype — but worth an explicit list of what stays static vs. moves to SharePoint (suggested split below).

**Accessibility:** nav links are `<a>` elements with no `href` — invisible to keyboard users. Tab buttons lack `role="tab"`/`aria-selected`. The booking modal has no focus trap. The search box does nothing.

**Cosmetic:** demo banner says `js/config.js` but files are flat. `jotform.js?v=4` cache-busting suggests manual versioning pain — a deploy pipeline solves this.

---

## Suggested roadmap

**Phase 1 — Make it work (this week)**
Rebuild the missing UI core (showPage/showToast/tabs), fix `window.AUTH`, remove the dead inline signIn, decide on hosting (recommendation: Azure SWA per DEPLOY.md — it gives you the Functions proxy for free).

**Phase 2 — Make it live (the actual product)**
Build the real graph.js: resolve site ID → fetch list items for Launches/Campaigns/Events → render into the existing card designs → list the document library in the Resources tab. Drive the countdown and homepage product news from the same data. This is what turns a beautiful mock-up into the tool.

**Phase 3 — Make it safe**
Azure Function proxy for Jotform (key in app settings, never in the browser). Migrate auth to MSAL with `ssoSilent` so staff genuinely never see a login screen on work devices. XSS pass on all interpolated attributes.

**Phase 4 — Make it lovely**
Hash-based routing (`#/launches`) so pages are bookmarkable and the back button works. Wire the search box. Accessibility pass. SessionStorage caching with a short TTL so page hops feel instant.

**What should stay static vs. go to SharePoint:** quick links, training pathways and survey widgets can stay hardcoded for v1 — they change rarely. Launches, campaigns, events, documents and the countdown must come from SharePoint or the hub will be stale within a fortnight and the team will stop trusting it.

---

## One strategic thought

The strongest thing about this hub is that marketing can update content in SharePoint without touching code — that's the pitch in DEPLOY.md and it's the right one. Every hardcoded section quietly undermines that promise. A useful discipline for the rest of the build: *if it will change more than once a quarter, it comes from a list.*
