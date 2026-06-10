# Hub App — What Changed (10 June 2026)

Prompts 1, 2, 3 and 5 from hub-app-prompts.md, executed. Drop these six files into your hub-app folder, **replacing** the old auth.js, graph.js, app.js, jotform.js and index.html. `ui.js` is new. config.js, styles.css, msal-browser.min.js and staticwebapp.config.json are unchanged.

## New: ui.js (Prompt 1 — the missing file)

Navigation (`showPage`), toasts (`showToast`), both tab systems (`switchTab`, `switchTrainingTab`), the Resources tab loader (`loadResourcesData`), plus the shared escaping helpers the whole app now uses: `escHtml`, `escAttr`, `safeUrl` (http/https only — kills `javascript:` URLs) and `safeCssUrl` (can't break out of `background-image:url(...)`). Loads first so everything else can rely on it.

## Rewritten: graph.js (Prompt 2 — real SharePoint integration)

The WordPress feed stays, now cached so the hero grid and news section share **one** request instead of two. New Graph layer: site ID resolution, `fetchListItems()` for your three lists, `fetchLibraryFiles()` for the document library, renderers for launches / campaigns / events / documents using your existing card styles, 5-minute sessionStorage cache, and a `loadSharePointData()` orchestrator using `Promise.allSettled` — one list failing doesn't break the others. Friendly errors name the misconfigured list. Design choice: if the Campaigns or Events lists are empty or fail, your hardcoded demo cards stay visible (an error note is added) rather than being wiped.

## Rewritten: auth.js (Prompt 3 — MSAL with silent SSO)

Hand-rolled PKCE replaced with MSAL. Order of attempts: redirect result → cached account → `ssoSilent` → sign-in page. The silent SSO uses a stored login hint, so from the **second** visit onwards staff on work devices are signed in with zero clicks. Public surface unchanged (`initAuth`, `getAccessToken`, `signIn`, `signOut`). **Bug fixed:** `window.AUTH` is now set explicitly — the booking form's account-manager pre-fill works for the first time.

## Hardened: app.js + jotform.js (Prompt 5 — security pass)

WordPress `post.link`/`post.image` now validated and escaped in href/CSS contexts; dates escaped. All Jotform answer values (company, time, visitor count, initials) escaped — they're user-submitted form data. `postMessage` handler now checks the message **origin** is jotform.com before trusting it. Visitor count parsed as a number. Bonus: the hardcoded "/ 22" free-days figure now counts the month's actual weekdays.

## Patched: index.html

MSAL script tag added to `<head>`; `ui.js` added to the script chain (config → ui → auth → graph → app → jotform, `?v=5`); the dead inline `signIn()` (which called a non-existent function) and the broken `showPage` wrapper removed in favour of a clean boot block.

## Still on the list (not in this pass)

The Jotform API key still lives client-side — run **Prompt 4** (Azure Function proxy) before going live, and don't paste the real key into config.js until then. Also outstanding: hosting decision (GitHub Pages vs Azure SWA), accessibility pass (Prompt 7), hash routing and search (Prompt 6).

## Verified

All five JS files pass `node --check`; escaping helpers unit-tested (`javascript:` URLs neutralised, HTML injection escaped, CSS url() breakout impossible). What I can't test from here is the live sign-in and Graph calls against your actual tenant — first run, check the browser console (F12) and the troubleshooting table in DEPLOY.md.
