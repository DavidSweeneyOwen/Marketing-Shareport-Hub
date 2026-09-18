/* Marketing Hub — 18 Sep 2026 batch harness.
 *
 * Static assertions plus a MEASURED run: config.js and graph.js are
 * loaded into a vm with a fake document, and the new code is driven
 * against fake SharePoint rows. The pattern is the one from 14/16 Sep —
 * assert on behaviour, not on the presence of a string, and re-assert
 * the fixes the previous rounds shipped so a wrong baseline shows up
 * here rather than on the live site.
 */
const fs = require('fs'), path = require('path'), vm = require('vm');
const D = __dirname;
let pass = 0, fail = 0;
const ok  = (c, m) => { if (c) { pass++; } else { fail++; console.log('  ✗ ' + m); } };
const read = f => fs.readFileSync(path.join(D, f), 'utf8');

const cfgSrc = read('config.js'), graphSrc = read('graph.js'),
      appSrc = read('app.js'),   html = read('index.html');

/* ── 1. Static ───────────────────────────────────────────── */
ok(/key: 'media', label: 'Media Portal', enabled: false,/.test(cfgSrc), 'media source is enabled:false');
ok(/key: 'sales', label: 'Sales & Marketing', enabled: false,/.test(cfgSrc), 'sales source is enabled:false');
ok(/key: 'portal', label: 'Product Portal',\n\s+site: 'https:\/\/checkfireltd\.sharepoint\.com\/sites\/CheckFireProductPortal',/.test(cfgSrc),
   'portal source untouched and still first');
ok(!/key: 'portal'[\s\S]{0,200}enabled: false/.test(cfgSrc), 'portal source is NOT disabled');
ok(/if \(src\.enabled === false\) return;/.test(graphSrc), 'the crawl honours enabled:false');
ok(/view: 'folders',/.test(cfgSrc), 'resources is in the folders view');
ok(/'Images for Resources'\]/.test(cfgSrc), 'the resources image folder is excluded from its own cards');
ok(/'FSE 2026': \{ start: '2026-04-28', end: '2026-04-30' \}/.test(cfgSrc), 'FSE 2026 has real dates');
ok(/config\.js\?v=24/.test(html) && /graph\.js\?v=27/.test(html) && /app\.js\?v=22/.test(html), 'version tags bumped');
ok(!/toggleResourcesBrowse\(this\)/.test(html), 'the Resources "Browse folders" / "Back to the library" button is gone');
ok(/togglePortalBrowse\(this\)/.test(graphSrc), 'the Product Portal keeps its own Browse folders button (it is drawn by graph.js)');
ok(/cardsFrom: \['Brand'\]/.test(cfgSrc), 'Brand is opened out into its sub-folders');
ok(/auth\.js\?v=2/.test(html) && /ui\.js\?v=3/.test(html) && /jotform\.js\?v=13/.test(html) && /ember\.js\?v=7/.test(html),
   'the untouched files keep their tags');

/* Fixes from earlier rounds that must still be here (the fix-4/fix-5
 * wrong-baseline trap). */
ok(/_ppLinksHtml\('\*'\)/.test(graphSrc), '11 Sep: the portal front page still shows every link');
ok(/GRAPH_MAX_INFLIGHT/.test(graphSrc), '14 Sep: the Graph concurrency gate is still there');
ok(/_libCacheWrite|_libCacheRead/.test(graphSrc), '16 Sep: the saved index is still there');
ok(/function _sameCat/.test(graphSrc), '16 Sep: case-insensitive category matching still there');
ok(/redirectUri/.test(read('config.js')) === false || true, '');
pass--;  // the line above is a no-op placeholder, don't count it

/* ── 2. Measured ─────────────────────────────────────────── */
// A fake DOM, small enough to reason about: elements remember their
// innerHTML, style and classes, and getElementById hands back the same
// object every time so a render can be inspected after the fact.
function makeDom() {
  const els = new Map();
  const mk = (id) => ({
    id, innerHTML: '', textContent: '', value: '',
    style: {}, _classes: new Set(),
    classList: { add(c){this._o._classes.add(c);}, remove(c){this._o._classes.delete(c);},
                 toggle(c,on){ on ? this._o._classes.add(c) : this._o._classes.delete(c); },
                 contains(c){ return this._o._classes.has(c); } },
    setAttribute(){}, getAttribute(){ return ''; }, querySelectorAll(){ return []; },
    querySelector(){ return null; }, closest(){ return null; }, appendChild(){},
  });
  return {
    getElementById(id) { if (!els.has(id)) { const e = mk(id); e.classList._o = e; els.set(id, e); } return els.get(id); },
    querySelectorAll() { return []; },
    querySelector() { return null; },
    addEventListener() {},
    createElement() { const e = mk('new'); e.classList._o = e; return e; },
    body: { appendChild() {} },
    readyState: 'complete',
    _els: els,
  };
}

const logs = [];
const doc = makeDom();
const ctx = {
  document: doc,
  console: { info: (...a) => logs.push(a.join(' ')), warn: (...a) => logs.push(a.join(' ')), log(){}, error(){} },
  setTimeout, clearTimeout, Promise, Map, Set, Date, JSON, Math, RegExp, Object, Array, String, Number, Boolean, isNaN,
  fetch: async () => { throw new Error('no network in the harness'); },
  localStorage: { getItem: () => null, setItem(){}, removeItem(){} },
  sessionStorage: { getItem: () => null, setItem(){}, removeItem(){} },
  location: { pathname: '/', origin: 'https://marketing.checkfire.co.uk' },
  history: { replaceState(){}, pushState(){}, state: null },
  navigator: { clipboard: null },
  Node: function(){},
};
ctx.addEventListener = function(){};
ctx.removeEventListener = function(){};
ctx.scrollTo = function(){};
ctx.requestAnimationFrame = function(cb){ return setTimeout(cb, 0); };
ctx.window = ctx;
ctx.globalThis = ctx;
vm.createContext(ctx);
// Loaded in the same order index.html loads them — escAttr/escHtml and
// the other helpers live in ui.js and app.js, and app.js's showPage is
// deliberately the one that wins, exactly as it does in the browser.
vm.runInContext(cfgSrc,  ctx, { filename: 'config.js' });
// ui.js is not part of this batch, so the harness runs with or without
// it: if a copy is sitting next to these files it is used as-is,
// otherwise the three helpers graph.js borrows from it are defined here.
// They are escaping functions — a stub that did LESS escaping would make
// the harness pass on markup the browser would choke on, so these are
// the real ones.
try {
  vm.runInContext(fs.readFileSync(path.join(D, 'ui.js'), 'utf8'), ctx, { filename: 'ui.js' });
} catch (_) {
  vm.runInContext(`
    function escHtml(s){ return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
    function escAttr(s){ return escHtml(s); }
    function safeCssUrl(u){ return String(u || '').replace(/["'()\\\\]/g, ''); }
  `, ctx, { filename: 'ui-stub.js' });
}
vm.runInContext(graphSrc, ctx, { filename: 'graph.js' });
vm.runInContext(appSrc,  ctx, { filename: 'app.js' });

// `const` bindings at the top of a script are NOT properties of the
// global object, so LIB has to be reached through the context itself.
const LIBv = vm.runInContext('LIB', ctx);

/* ── 2a. Events: upcoming is a DATE now ──────────────────── */
const today = new Date('2026-09-18');
let w = ctx._evWhen('FSE 2026', 2026, today);
ok(w.upcoming === false, 'FSE 2026 reads as PREVIOUS on 18 Sep 2026 (the reported bug)');
ok(w.dated === true, 'FSE 2026 is dated, not assumed');
ok(w.label === '28–30 Apr 2026', 'FSE 2026 shows its dates, not the bare year (got "' + w.label + '")');

ok(ctx._evWhen('FSE 2026', 2026, new Date('2026-04-15')).upcoming === true,  'FSE 2026 WAS upcoming in April');
ok(ctx._evWhen('FSE 2026', 2026, new Date('2026-04-30')).upcoming === true,  'an event is upcoming on its last day');
ok(ctx._evWhen('FSE 2026', 2026, new Date('2026-05-01')).upcoming === false, 'and previous the day after');
ok(ctx._evWhen('FSE2026',  2026, today).dated === true, 'the date row matches on letters and digits ("FSE2026")');

ok(ctx._evWhen('FSE 2027', 2027, today).upcoming === true,  'FSE 2027 is upcoming');
ok(ctx._evWhen('FSE 2025', 2025, today).upcoming === false, 'FSE 2025 is previous');
const before = logs.length;
const undated = ctx._evWhen('Open Day 2026', 2026, today);
ok(undated.upcoming === false, 'an undated event this year falls back to previous');
ok(logs.length > before && /Open Day 2026/.test(logs[logs.length - 1]),
   'and the console names it with the row to add');
ok(ctx._evWhen('Showroom refresh', null, today).upcoming === true, 'a folder with no year is still treated as ongoing');

// Sorting: soonest first while upcoming, most recent first once past.
const rows = [
  { name: 'FSE 2027', year: 2027 }, { name: 'FSE 2026', year: 2026 }, { name: 'FSE 2025', year: 2025 },
].map(r => Object.assign(r, ctx._evWhen(r.name, r.year, today)));
const order = (a, b) => (a.upcoming === b.upcoming ? 0 : a.upcoming ? -1 : 1) || (a.upcoming ? a.at - b.at : b.at - a.at);
const sorted = rows.slice().sort(order).map(r => r.name);
ok(JSON.stringify(sorted) === JSON.stringify(['FSE 2027', 'FSE 2026', 'FSE 2025']),
   'events sort upcoming-then-most-recent (got ' + sorted.join(', ') + ')');

// Sibling-year artwork
ok(ctx._evBaseName('FSE 2027') === ctx._evBaseName('FSE 2026'), 'FSE 2027 and FSE 2026 are the same event for artwork');
ok(ctx._evBaseName("FSE '26") === ctx._evBaseName('FSE 2026'), "and so is FSE '26");
ok(ctx._evBaseName('Open Day 2026') !== ctx._evBaseName('FSE 2026'), 'but a different event is not');

/* ── 2b. Resources folder cards ──────────────────────────── */
// The real shape of MarketingHub ▸ Documents, read from SharePoint:
// Brand holds the three folders David wants as cards, and the rest of
// the root is either excluded or empty.
const files = [
  { id: 'a', name: 'CheckFire Brand Guidelines.pdf', _driveId: 'd', _path: ['Brand', 'BRAND GUIDELINES'],      lastModifiedDateTime: '2026-08-13T00:00:00Z' },
  { id: 'b', name: 'logo-primary.png',               _driveId: 'd', _path: ['Brand', 'BRAND GUIDELINES'],      lastModifiedDateTime: '2026-08-13T00:00:00Z' },
  { id: 'c', name: 'Customer deck.pptx',             _driveId: 'd', _path: ['Brand', 'CUSTOMER PRESENTATIONS'], lastModifiedDateTime: '2026-09-16T00:00:00Z' },
  { id: 'd', name: 'Toolkit.zip',                    _driveId: 'd', _path: ['Brand', 'MARKETING TOOLKIT'],     lastModifiedDateTime: '2026-08-13T00:00:00Z' },
  { id: 'e', name: 'Loose note.docx',                _driveId: 'd', _path: ['Brand'],                          lastModifiedDateTime: '2026-07-01T00:00:00Z' },
];
LIBv.resources = Object.assign(
  { files: [], loaded: true, driveId: 'd', tag: 'all', cat: 'all', q: '' },
  { files: ctx._libDecorate('resources', files), loaded: true, driveId: 'd', tag: 'all', cat: 'all', q: '' });

const folders = ctx._libFolderRows('resources');
const names = folders.map(f => f.label);
ok(!!folders.find(f => f.label === 'BRAND GUIDELINES'), 'BRAND GUIDELINES is a card');
ok(!!folders.find(f => f.label === 'CUSTOMER PRESENTATIONS'), 'CUSTOMER PRESENTATIONS is a card');
ok(!!folders.find(f => f.label === 'MARKETING TOOLKIT'), 'MARKETING TOOLKIT is a card');
ok(!names.some(n => n === 'Brand' && folders.length === 1), 'the container is not the card any more');
const brand = folders.find(f => f.label === 'BRAND GUIDELINES');
ok(!!brand && brand.n === 2, 'the Brand Guidelines card counts both of its files');
ok(brand && brand.img && brand.img.name === 'logo-primary.png', 'a picture already in the folder becomes the card artwork');
ok(!!folders.find(f => f.label === 'Brand' && f.n === 1), 'a file loose inside the opened-out folder keeps its own card — nothing falls off the page');
ok(folders[0].label === 'CUSTOMER PRESENTATIONS', 'the most recently updated folder leads (got ' + folders[0].label + ')');
ok(ctx._libCardLabel('resources', ['Campaigns', 'LFX']) === 'Campaigns', 'a folder NOT named in cardsFrom is still one card');

ctx.renderLibraryFolders('resources');
const host = doc.getElementById('res-index');
ok(/px-grid/.test(host.innerHTML) && /px-card/.test(host.innerHTML), 'the cards render on the .px-card grid, like the other pages');
ok((host.innerHTML.match(/px-card /g) || []).length === folders.length, 'one card per folder');
ok(/libOpenFolder\('resources',0\)/.test(host.innerHTML), 'a card opens by INDEX, never by folder name');
ok(!/onclick="libOpenFolder\('resources','/.test(host.innerHTML), 'no folder name ever reaches a handler (apostrophe trap)');
ok(/lib-results-resources/.test(host.innerHTML) && /style="display:none"/.test(host.innerHTML),
   'the results list starts hidden behind the cards');

ctx.libOpenFolder('resources', folders.indexOf(brand));
ok(LIBv.resources.card === 'BRAND GUIDELINES', 'opening a card filters to that folder');
ok(LIBv.resources.cat === 'all', 'and does it on the CARD, not the category (every file here reads _cat "Brand")');
ok(doc.getElementById('lib-front-resources').style.display === 'none', 'the cards step aside');
ok(doc.getElementById('lib-results-resources').style.display === '', 'the documents come forward');
ok(/libFoldersBack\('resources'\)/.test(doc.getElementById('lib-head-resources').innerHTML), 'with a way back');
const shown = doc.getElementById('lib-results-resources').innerHTML;
ok(/Brand Guidelines/.test(shown), 'the folder shows its own files');
ok(!/Customer deck/.test(shown), 'and only its own files');
ok(!/Loose note/.test(shown), 'and not the ones loose in the container above it');
ok(/BRAND GUIDELINES<span>2<\/span>/.test(shown.replace(/\s+/g, '')) || /BRAND GUIDELINES/.test(shown),
   'the group heading names the card, not the container');

ctx.libFoldersBack('resources');
ok(!LIBv.resources.card, 'back clears the filter');
ok(doc.getElementById('lib-front-resources').style.display === '', 'and the cards come back');
ok(doc.getElementById('lib-results-resources').style.display === 'none', 'with the results hidden again');

// Search from the front opens the results across every folder.
ctx.libSearch('resources', 'deck');
const done = new Promise(r => setTimeout(r, 200));

/* ── 2c. Image matching floor ────────────────────────────── */
const pool = [
  { holder: 'Brand & guidelines', name: 'brand.png', keys: [ctx._slugKey('Brand & guidelines'), ctx._slugKey('brand')],
    words: [ctx._portalWords('Brand & guidelines'), ctx._portalWords('brand')], url: 'BRAND' },
  { holder: 'CO2 extinguisher', name: 'co2.png', keys: [ctx._slugKey('CO2 extinguisher'), ctx._slugKey('co2')],
    words: [ctx._portalWords('CO2 extinguisher'), ctx._portalWords('co2')], url: 'CO2' },
];
const picked = ctx.assignImagesToCards(pool, [{ key: 'brand', label: 'Brand & guidelines' }, { key: 'pres', label: 'Presentations' }]);
ok(picked.get('brand') === 'BRAND', 'a folder gets the picture named after it');
ok(picked.get('pres') === undefined, 'and a folder with no match keeps its initials rather than taking a stray picture');

/* ── 2c-ii. The Product Portal's own artwork still places ──
 * _imagePool was lifted out of fetchPortalImages this round, so the
 * matching the portal cards depend on is re-asserted here against the
 * REAL sub-folder names marketing built, including the near-miss trap
 * from deck 8: a folder called "Certificates" must not take the
 * "Certificates & Declarations" card.
 */
const ppRow = (holder) => ({
  holder, name: holder + '.png', folder: holder,
  keys: [ctx._slugKey(holder), ctx._slugKey(holder)],
  words: [ctx._portalWords(holder), ctx._portalWords(holder)],
  url: holder,
});
const ppPool = ['Main Product portal image', 'Datasheets & MSDS', 'Certificates & Declarations',
                'Product Information Files', 'Certificates'].map(ppRow);
const ppSecs = [
  { key: 'ds',   label: 'Datasheets & MSDS' },
  { key: 'cert', label: 'Certificates & Declarations' },
  { key: 'pif',  label: 'Product Information Files' },
];
const ppPicked = ctx.assignPortalImages(ppPool, ppSecs);
ok(ppPicked.get('ds')   === 'Datasheets & MSDS',        'the datasheets card keeps its picture');
ok(ppPicked.get('cert') === 'Certificates & Declarations', 'and the certificates card is not stolen by "Certificates"');
ok(ppPicked.get('pif')  === 'Product Information Files', 'and the PIF card keeps its picture');
ok(ctx.mainPortalImage(ppPool) === 'Main Product portal image', 'the lead image is still picked out by its folder name');
ok(ppPicked.get('ds') !== ppPicked.get('cert'), 'one picture, one card');

/* ── 2d. app.js: every tab opens on its own front ─────────── */
ok(/function resetPageView\(id\)/.test(appSrc), 'resetPageView exists');
ok(/async function showPage\(id, idx\) \{\n  resetPageView\(id\);/.test(appSrc), 'showPage resets the page it is opening, first thing');
ok(/closeLaunchDetail\(\);/.test(appSrc) && /closeCampaignDetail\(\);/.test(appSrc) && /closeEventFolder\(\);/.test(appSrc),
   'launches, campaigns and events are all put back');
ok(/lib && lib\.product && lib\.product\.loaded/.test(appSrc), 'the portal is only reset once it has loaded (skeleton trap)');
ok(/lib && lib\.resources && lib\.resources\.loaded/.test(appSrc), 'and so is Resources');
ok(/cardsBackLabel/.test(graphSrc) && /cardsBackLabel: 'All folders'/.test(cfgSrc),
   'the way back out of a card is not called "Back to the library"');
ok(/if \(d && d\.deep && d\.page && typeof showPage === 'function'\) showPage\(d\.page\);/.test(appSrc),
   'a deep history entry puts its page back before replaying');
ok(/'libOpenFolder', 'libFoldersBack'/.test(appSrc), 'the folder cards are in the history router');
ok(/if \(fromId === 'reader' \|\| fromId === 'search'\) return;/.test(appSrc),
   'coming back out of the reader leaves the page it returns to alone');
ok(/if \(_libCfg\('resources'\)\.view === 'folders'\)[\s\S]{0,240}libOpenFolder\('resources', i\)/.test(graphSrc),
   'the site search opens a Resources folder card rather than filtering a hidden list');
// The wrapping is still install-time, so no call site had to change.
ok(/PAGE_FNS\.forEach\(function \(n\) \{ wrap\(n, false\); \}\);/.test(appSrc), '17 Sep: the router still wraps at install time');

done.then(() => {
  ok(doc.getElementById('lib-front-resources').style.display === 'none', 'typing in the search box opens the results over the cards');
  ok(/Customer deck/.test(doc.getElementById('lib-results-resources').innerHTML), 'and searches every folder');
  ctx.libSearch('resources', '');
  setTimeout(() => {
    ok(doc.getElementById('lib-front-resources').style.display === '', 'clearing the box puts the cards back');
    console.log(`\n${pass}/${pass + fail} checks passed.`);
    process.exit(fail ? 1 : 0);
  }, 200);
});
