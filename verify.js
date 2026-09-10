/* Verification harness — 10 Sep 2026 batch (deck 7 + Lowri).
 *
 * House rule from fix 5: assert on PARSED STATEMENTS, never on a
 * substring that a comment could satisfy. Every source assertion below
 * strips comments first. The landing-image checks go further and
 * actually RUN the shipped functions against the real folder and file
 * names pulled out of SharePoint today.
 */
const fs = require('fs');
let pass = 0, fail = 0;
const ok  = (n) => { pass++; console.log('  ok   ' + n); };
const bad = (n, d) => { fail++; console.log('  FAIL ' + n + (d ? '  → ' + d : '')); };
const is  = (n, a, b) => (a === b ? ok(n) : bad(n, `got ${JSON.stringify(a)} want ${JSON.stringify(b)}`));
const yes = (n, c, d) => (c ? ok(n) : bad(n, d));

const raw = {
  graph:  fs.readFileSync('graph.js', 'utf8'),
  app:    fs.readFileSync('app.js', 'utf8'),
  config: fs.readFileSync('config.js', 'utf8'),
  html:   fs.readFileSync('index.html', 'utf8'),
};

// Strip // and /* */ comments so no assertion can be satisfied by prose.
function decomment(s) {
  return s.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:])\/\/.*$/gm, '$1');
}
const src = Object.fromEntries(Object.entries(raw).map(([k, v]) => [k, decomment(v)]));
// HTML comments too.
src.html = raw.html.replace(/<!--[\s\S]*?-->/g, '');

console.log('\n1. Landing images — the duplicated banner');

// Pull the two functions out and run them for real.
function grab(name) {
  const i = raw.graph.indexOf('function ' + name + '(');
  if (i < 0) throw new Error('missing function ' + name);
  let d = 0, started = false;
  for (let j = i; j < raw.graph.length; j++) {
    const c = raw.graph[j];
    if (c === '{') { d++; started = true; }
    else if (c === '}') { d--; if (started && d === 0) return raw.graph.slice(i, j + 1); }
  }
  throw new Error('unbalanced ' + name);
}
const HUB_CONFIG = {
  landingImages: {
    folder: 'Images for Landing Pages',
    noiseWords: ['landing','page','pages','image','images','hero','banner',
                 'main','cover','final','new','copy','v1','v2','checkfire','cf'],
    minWordMatch: 1,
  },
};
const _slugKey = s => String(s || '').toLowerCase().replace(/\.[a-z0-9]+$/i, '').replace(/[^a-z0-9]+/g, '');
// The extracted functions get this console, not the global one, so the
// spy below has to listen here.
const LOG = [];
const spyConsole = { info: (...a) => LOG.push(['info', a.join(' ')]),
                     warn: (...a) => LOG.push(['warn', a.join(' ')]) };
const sandbox = { HUB_CONFIG, _slugKey, URL, console: spyConsole, Set, Map };
const names = ['_landingWords','_landingPageWords','_landingScore','assignLandingImages','matchLandingImage'];
const body  = names.map(grab).join('\n') + '\nmodule.exports={' + names.join(',') + '};';
const mod   = new Function('HUB_CONFIG','_slugKey','URL','console','module', body);
const M     = { exports: {} };
mod(sandbox.HUB_CONFIG, sandbox._slugKey, URL, sandbox.console, M);
const { _landingWords, assignLandingImages, matchLandingImage } = M.exports;

// The real state of SharePoint on 10 Sep 2026, read with the Microsoft
// 365 connector: `Images for landing pages` has NO Flat-Pack Tubular
// folder, and the Commander folder holds the Tubular banner.
const img = (file, folder) => ({
  keys: [_slugKey(file), _slugKey(folder), _slugKey(folder + ' ' + file)].filter(Boolean),
  words: [_landingWords(file), _landingWords(folder)],
  url: 'URL::' + folder + '/' + file,
  name: file, folder,
});
const images = [
  img('Black-Tubular-HP-Banner-1-1707x2048.png', 'Flat-Pack Commander Stand Landing Page'),
  img('bridgehill-10 (1).png',                   'Bridgehill Fire Blankets Landing Page'),
  img('CheckFire brochure.png',                  'Brochure Landing Page'),
  img('Fire Equipment Supplier Landing Page Image.png', 'Fire Equipment Supplier Landing Page'),
];
const page = (title, slug) => ({ title, link: 'https://www.checkfire.co.uk/' + slug + '/' });
const pages = [
  page('Flat-Pack Tubular Stand',   'flat-pack-tubular-stand'),
  page('Bridgehill Fire Blankets',  'bridgehill-fire-blankets'),
  page('CheckFire 2026 Brochure',   'checkfire-2026-brochure'),
  page('Flat-Pack Commander Stand', 'flat-pack-commander-stand'),
  page('Fire Equipment Suppliers',  'fire-equipment-suppliers'),
];

// THE regression: the old scorer gave this one image to both Flat-Pack
// pages. Prove no image is handed out twice.
const got  = assignLandingImages(images, pages);
const urls = [...got.values()];
is('no image is used by two pages', urls.length, new Set(urls).size);
yes('Bridgehill still matches its own folder',
    got.get(1) === 'URL::Bridgehill Fire Blankets Landing Page/bridgehill-10 (1).png', got.get(1));
yes('Brochure still matches its own folder',
    got.get(2) === 'URL::Brochure Landing Page/CheckFire brochure.png', got.get(2));
yes('Fire Equipment Suppliers still matches (the 1 Sep fix holds)',
    got.get(4) === 'URL::Fire Equipment Supplier Landing Page/Fire Equipment Supplier Landing Page Image.png', got.get(4));
const tub = got.get(0), com = got.get(3);
yes('exactly one of the two Flat-Pack pages gets the single Flat-Pack image',
    (tub ? 1 : 0) + (com ? 1 : 0) === 1, `tubular=${tub} commander=${com}`);
// The image lives in the Commander folder, so Commander is who it goes
// to — the folder is the thing marketing organise by. The Tubular page
// getting NOTHING is the correct, honest outcome while its folder does
// not exist; the console must say so, and must also flag that the file
// in Commander's folder is named for Tubular. Code cannot fix filing.
yes('the image goes to the page whose FOLDER it sits in', !!com && !tub, `tubular=${tub} commander=${com}`);
{
  LOG.length = 0;
  assignLandingImages(images, pages);
  const said = LOG.slice();
  yes('console names the folder marketing must create',
      said.some(([k, m]) => k === 'info' && m.includes('Flat-Pack Tubular Stand Landing Page')),
      JSON.stringify(said.filter(x => x[0] === 'info')));
  yes('console flags the misfiled banner by name',
      said.some(([k, m]) => k === 'warn' && m.includes('Black-Tubular-HP-Banner') && m.includes('Flat-Pack Tubular Stand')),
      JSON.stringify(said.filter(x => x[0] === 'warn')));
}

// Once marketing add the missing folder, both must resolve, separately.
const images2 = images.concat([img('Commander-Stand-Banner.png', 'Flat-Pack Commander Stand Landing Page')]);
const images3 = [
  img('Black-Tubular-HP-Banner-1-1707x2048.png', 'Flat-Pack Tubular Stand Landing Page'),
  img('Commander-Stand-Banner.png',              'Flat-Pack Commander Stand Landing Page'),
].concat(images.slice(1));
const got3 = assignLandingImages(images3, pages);
yes('with both folders present, Tubular gets the tubular art',
    String(got3.get(0)).includes('Tubular Stand Landing Page'), got3.get(0));
yes('with both folders present, Commander gets the commander art',
    String(got3.get(3)).includes('Commander-Stand-Banner'), got3.get(3));
is('with both folders present nothing is shared', [...got3.values()].length, new Set([...got3.values()]).size);
yes('matchLandingImage still returns a string for a single page',
    typeof matchLandingImage(images, pages[1]) === 'string');
void images2;

console.log('\n2. Carousel order');
yes('_orderLandingPages exists in app.js', /function\s+_orderLandingPages\s*\(/.test(src.app));
yes('loadLandingPages runs pages through it', /_orderLandingPages\(await fetchWordPressPages\(\)\)/.test(src.app));
yes('assignLandingImages is what app.js calls', /assignLandingImages\(images,\s*pages\)/.test(src.app));
yes('the old per-page loop is gone', !/const url = matchLandingImage\(images, p\);/.test(src.app));
for (const p of ['Flat-Pack Tubular Stand','Bridgehill Fire Blankets','CheckFire 2026 Brochure','Flat-Pack Commander Stand'])
  yes(`pinned list carries "${p}"`, src.config.includes(`'${p}'`));

console.log('\n3. Deck 7 deletions');
yes('.px-lead-sub is no longer rendered', !/class="px-lead-sub"/.test(src.graph));
yes('launch lead passes no sub', !/sub: lead\.Description \|\| lead\.Summary \|\| '',/.test(src.graph));
yes('campaign lead passes no sub (the bare "Brand")',
    !/sub: lead\.Description \|\| lead\.Summary \|\| lead\.CampaignType/.test(src.graph));

console.log('\n4. Trade & events order');
const evHtml = src.graph.match(/host\.innerHTML = `\s*<section id="ev-training">[\s\S]{0,200}?`;/);
yes('training section is emitted before the folder categories', !!evHtml);
if (evHtml) {
  const t = evHtml[0].indexOf('ev-training'), c = evHtml[0].indexOf('cats.map(section)');
  yes('…and in that order in the template', t >= 0 && c > t, `training@${t} cats@${c}`);
}
yes('heading reads "Training sessions"', src.graph.includes('>Training sessions</h2>'));
yes('no "Training &amp; sessions" left', !src.graph.includes('Training &amp; sessions'));
yes('the hover menu leads with training too',
    /const rows = \[_navRow\('Training sessions'/.test(src.graph));

console.log('\n5. Product Portal — the product team\'s own lists');
for (const f of ['fetchPortalOverrides','ppSections','ppLinks','renderPortalLinks','_ppScopeTypeChips'])
  yes(`${f}() defined`, new RegExp('function\\s+' + f + '\\s*\\(').test(src.graph));
yes('loadProductPortal awaits the overrides', /await fetchPortalOverrides\(\);/.test(src.graph));
yes('loadProductPortal renders the links band', /renderPortalLinks\(\);/.test(src.graph));
yes('renderPortalSections reads ppSections()', /const secs\s*=\s*ppSections\(\)/.test(src.graph));
yes('the nav menu reads ppSections() too', /const secs = ppSections\(\);/.test(src.graph));
yes('config declares the two list names', /portalSections:\s*'Portal Sections'/.test(src.config) && /portalLinks:\s*'Portal Links'/.test(src.config));
yes('#pp-links host exists in the HTML', /id="pp-links"/.test(src.html));
yes('.pp-link CSS exists', /\.pp-link\{/.test(src.html));

console.log('\n6. Lowri\'s specific asks');
yes('feedback form URL is wired', /feedbackUrl:\s*'https:\/\/form\.jotform\.com\/261374017099056'/.test(src.config));
for (const l of ['CheckFire Product Information File','PJ Fire Product Information File','Sample Request Sheet','New Product Request Sheet','Product Feedback Form'])
  yes(`link present: ${l}`, src.config.includes(l));
yes('capital D — Certificates & Declarations', /label:'Certificates & Declarations'/.test(src.config));
yes('capital I — Manuals & Instructions',      /label:'Manuals & Instructions'/.test(src.config));
yes('capital T — Product Training',            /label:'Product Training'/.test(src.config));
yes('tile counts off for the portal', /showTileCounts:\s*false/.test(src.config));
yes('recently-updated row off for the portal', /showRecent:\s*false/.test(src.config));
yes('type chips scoped to the open section', /scopeTypesToSection:\s*true/.test(src.config));
yes('the tile count is conditional in the renderer', /\$\{showCounts \? `<span class="lib-tile-n">/.test(src.graph));
yes('recent row is suppressed by config, not deleted', /const wantRecent = key === 'product' \? ppCfg\.showRecent !== false : true;/.test(src.graph));
yes('resources page keeps its recent row', /key === 'product' \? ppCfg\.showRecent !== false : true/.test(src.graph));

console.log('\n7. Fail-safe behaviour');
yes('missing Portal Sections list falls back to config', /if \(!ov \|\| !ov\.length\) return base;/.test(src.graph));
yes('missing Portal Links list falls back to config', /return \(HUB_CONFIG\.productPortal && HUB_CONFIG\.productPortal\.links\) \|\| \[\];/.test(src.graph));
yes('sections not named by the list are still shown', /base\.forEach\(s => \{ if \(!claimed\.has\(s\.key\)\) out\.push\(s\); \}\);/.test(src.graph));
yes('an empty override never blanks the page', /return out\.length \? out : base;/.test(src.graph));
yes('list reads are wrapped in try/catch', (src.graph.match(/catch \(e\) \{ console\.info\('\[Portal\] Portal .* not read/g) || []).length === 2);

console.log('\n7b. The lists are on the Product Portal site, not MarketingHub');
yes('fetchListItemsOn() takes a site', /async function fetchListItemsOn\(siteUrl, listName\)/.test(src.graph));
yes('it resolves the site it was given, not getSiteId()', /const siteId = await resolveSiteId\(siteUrl\);/.test(src.graph));
yes('_fetchPortalList tries both sites', /const sites = \[HUB_CONFIG\.productPortalSite, HUB_CONFIG\.sharepointSite\]/.test(src.graph));
yes('Product Portal is tried first', (() => {
  const m = src.graph.match(/const sites = \[([^\]]+)\]/);
  return !!m && m[1].indexOf('productPortalSite') < m[1].indexOf('sharepointSite');
})());
yes('both override reads go through _fetchPortalList',
    (src.graph.match(/await _fetchPortalList\(names\.portal/g) || []).length === 2);
yes('neither still calls the MarketingHub-only fetchListItems',
    !/fetchListItems\(names\.portal/.test(src.graph));
yes('cache key includes the site so the two never collide',
    /const cacheKey = 'list_' \+ siteUrl \+ '::' \+ listName;/.test(src.graph));
yes('_ppField() tolerates alternative column names', /function _ppField\(row, names\)/.test(src.graph));
yes('…including a space-and-case-insensitive fallback',
    /k\.toLowerCase\(\)\.replace\(\/\[\^a-z0-9\]\/g, ''\)/.test(src.graph));
yes('rows present but unusable are reported, not swallowed',
    (src.graph.match(/console\.warn\(`\[Portal\] "\$\{names\.portal/g) || []).length === 2);
yes('the warning names the columns it actually saw',
    (src.graph.match(/Columns seen: \$\{Object\.keys\(rows\[0\] \|\| \{\}\)\.join\(', '\)\}/g) || []).length === 2);
yes('absent Show column means visible', /if \(v === undefined \|\| v === null \|\| v === ''\) return true;/.test(src.graph));

console.log('\n8. Josh avatar');
yes('avatar is inlined as a data URI', /background-image:url\('data:image\/webp;base64,/.test(src.html));
yes('applied to .ember-av so per-message avatars get it', /\.ember-av\{[\s\S]*?background-image:url\('data:image\/webp/.test(src.html));
yes('the letter is hidden, not removed', /id="ember-av">J<\/span>/.test(src.html));

console.log('\n9. Nothing else moved');
yes('config.js still parses', (() => { try { new Function(raw.config); return true; } catch (_) { return false; } })());
yes('index.html div balance', (raw.html.match(/<div[\s>]/g) || []).length === (raw.html.match(/<\/div>/g) || []).length);
yes('index.html section balance', (raw.html.match(/<section[\s>]/g) || []).length === (raw.html.match(/<\/section>/g) || []).length);

// Every element id the changed code touches must exist in the HTML.
for (const id of ['pp-links','pp-sections','pp-section-head','pp-index','pp-upcoming','pp-feedback','ember-av','home-pages-track','home-pages-section'])
  yes(`#${id} present in index.html`, src.html.includes(`id="${id}"`));
// #ev-training is created by loadTradeEvents(), not written in the HTML.
yes('#ev-training is emitted by graph.js', src.graph.includes('<section id="ev-training">'));
yes('renderTrainingList() targets it', /getElementById\('ev-training'\)/.test(src.graph));

console.log(`\n${pass} passing, ${fail} failing\n`);
process.exit(fail ? 1 : 0);
