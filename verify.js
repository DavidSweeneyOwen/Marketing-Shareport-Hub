// Harness for the 16 Sep 2026 batch — deck 8: Product Portal artwork
// and option A.
//
//   node "UPLOAD TO GITHUB - 16 Sep 2026 (deck 8 - portal images + option A)\verify.js"
//
// Static assertions run against a DECOMMENTED copy of the shipped
// source, so a comment can never satisfy one (the fix-5 rule). The
// matcher is then EXTRACTED and run against the exact folder and file
// names marketing actually created in SharePoint on 16 Sep, because
// "it should match" is not the same as "it matched".

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const DIR = __dirname;
let pass = 0, fail = 0;
const ok  = (n) => { pass++; console.log('  PASS  ' + n); };
const bad = (n, d) => { fail++; console.log('  FAIL  ' + n + (d ? '  — ' + d : '')); };
const is  = (n, a, b) => (a === b ? ok(n) : bad(n, `expected ${b}, got ${a}`));
const yes = (n, c, d) => (c ? ok(n) : bad(n, d));

function decomment(src) {
  const noBlock = src.replace(/\/\*[\s\S]*?\*\//g, '');
  return noBlock.split('\n').map(l => l.replace(/^\s*\/\/.*$/, '')).join('\n');
}

const RAW  = fs.readFileSync(path.join(DIR, 'graph.js'), 'utf8');
const CODE = decomment(RAW);
const CFG  = fs.readFileSync(path.join(DIR, 'config.js'), 'utf8');
const HTML = fs.readFileSync(path.join(DIR, 'index.html'), 'utf8');
const HTMLC = decomment(HTML.replace(/<!--[\s\S]*?-->/g, ''));

// ── static ─────────────────────────────────────────────────────
console.log('\nStatic');
yes('index.html asks for graph.js v23', /graph\.js\?v=23/.test(HTML));
yes('index.html asks for config.js v20', /config\.js\?v=20/.test(HTML));
yes('the lead host exists in the page', /id="pp-lead"/.test(HTML));
yes('the old .ph header is gone from the portal',
    !/ph-title">Product portal</.test(HTMLC));
yes('there is exactly ONE pp-browse-btn',
    (RAW.match(/id="pp-browse-btn"/g) || []).length + (HTML.match(/id="pp-browse-btn"/g) || []).length === 1);
yes('the sub-type chip style is defined', /\.pp-subchip\{/.test(HTMLC));
yes('the dark feedback panel is defined', /\.pp-fb\.dark\{/.test(HTMLC));

console.log('\nStatic — option A in the renderer');
yes('the front renders the shared card grid', /class="px-grid"/.test(CODE));
yes('the front renders the shared rail', /class="px-rail pp-rail"/.test(CODE));
yes('section cards carry an image host', /id="pp-img-\$\{escAttr\(sec\.key\)\}"/.test(CODE));
yes('the lead carries an image host', /id="pp-lead-media"/.test(CODE));
yes('ppOpenSection takes a sub-type', /function ppOpenSection\(i,\s*subIndex\)/.test(CODE));
yes('a sub-type cannot widen a section', /subIndex\s*<\s*bandCats\.length/.test(CODE));
// A SharePoint label with an apostrophe in it must never reach an
// onclick — the attribute is decoded before the JS is parsed, so a
// label there is a broken handler waiting to happen.
yes('no category label is interpolated into a handler',
    !/ppOpenSection\(\$\{i\},'/.test(CODE));
yes('the plumbing note is off the front',
    !/documents from \$\{_ppSourceCount\(state\)\} SharePoint sources/.test(CODE));
yes('the feedback panel is the dark one', /class="pp-fb dark"/.test(CODE));
yes('artwork never blocks the documents', /imgs\.then\(paintPortalImages\)/.test(CODE));

console.log('\nStatic — what must NOT have been lost');
// 11 Sep: Lowri's links had nowhere to land. That fix is upstream of
// everything here and has been broken by a wrong baseline once already.
yes('11 Sep links fix is still present', /_ppLinksHtml\('\*'\)/.test(CODE));
yes("11 Sep front-mode is still present", /const\s+front\s*=\s*sectionKey\s*===\s*'\*'/.test(CODE));
// 14 Sep: the speed work.
yes('14 Sep gate is still present', /const\s+GRAPH_MAX_INFLIGHT\s*=\s*\d+/.test(CODE));
yes('14 Sep parallel roots still present', /Promise\.all\(src\.roots\.map/.test(CODE));
// 9 Sep, Lowri.
yes('scopeTypesToSection still honoured', /scopeTypesToSection\s*===\s*false/.test(CODE));
yes('showTileCounts still false in config', /showTileCounts:\s*false/.test(CFG));
yes('showRecent still false in config', /showRecent:\s*false/.test(CFG));

console.log('\nStatic — deck 8 config');
yes('Images for Product Portal is off Resources',
    /excludeFolders:\s*\[[^\]]*Images for Product Portal/.test(CFG));
yes('Images for Landing Pages is STILL off Resources',
    /excludeFolders:\s*\[[^\]]*Images for Landing Pages/.test(CFG));

// ── run config.js for real ─────────────────────────────────────
const cfgBox = { console: { info() {}, warn() {}, log() {} } };
vm.createContext(cfgBox);
vm.runInContext(CFG + '\n;this.__cfg = HUB_CONFIG;', cfgBox);
const HUB = cfgBox.__cfg;

yes('quickLinks keeps Media Portal, Website, LinkedIn',
    ['Media Portal', 'Website', 'CF LinkedIn'].every(l => HUB.quickLinks.some(q => q.label === l)));
yes('quickLinks adds the two Smartsheet forms',
    HUB.quickLinks.filter(q => /smartsheet\.com/.test(q.url)).length === 2);
yes('no FPS placeholder URL shipped live',
    !HUB.quickLinks.some(q => /PASTE THE/.test(q.url)));
yes('portalImages points at the real folder',
    HUB.portalImages && HUB.portalImages.folder === 'Images for Product Portal');
yes('every section has an eyebrow',
    HUB.productPortal.sections.every(s => !!s.eyebrow));

// ── extract the matcher and run it ─────────────────────────────
function grab(name) {
  const i = CODE.indexOf('function ' + name + '(');
  if (i < 0) throw new Error('could not find ' + name);
  let d = 0, j = CODE.indexOf('{', i);
  for (let k = j; k < CODE.length; k++) {
    if (CODE[k] === '{') d++;
    else if (CODE[k] === '}') { d--; if (!d) return CODE.slice(i, k + 1); }
  }
  throw new Error('unbalanced ' + name);
}

const names = ['_slugKey', '_portalWords', '_portalIsMain', '_portalScore',
               'assignPortalImages', 'mainPortalImage', '_ppSubTypes'];
const box = {
  console: { info() {}, warn() {}, log() {} },
  HUB_CONFIG: HUB, Math, JSON, Set, Map, Array, String, Object,
};
vm.createContext(box);
vm.runInContext(names.map(grab).join('\n') +
  '\n;this.__api = {' + names.join(',') + '};', box);
const A = box.__api;

// The real thing, as surveyed in SharePoint on 16 Sep 2026:
//   Documents ▸ Images for Product Portal ▸
//     Main Product portal image ▸ main product portal image.png
//     Datasheets & MSDS         ▸ Data sheet and MSDS image for product portal.png
//     Certificates & Declarations ▸ Certificates & Declarations image.png
//     Product Information Files ▸ Product information files image.png
const mk = (holder, name) => ({
  holder, name, folder: 'Images for Product Portal/' + holder,
  keys:  [A._slugKey(holder), A._slugKey(name.replace(/\.[a-z0-9]+$/i, ''))].filter(Boolean),
  words: [A._portalWords(holder), A._portalWords(name)],
  url:   'URL::' + holder,
});

const IMAGES = [
  mk('Main Product portal image',   'main product portal image.png'),
  mk('Datasheets & MSDS',           'Data sheet and MSDS image for product portal.png'),
  mk('Certificates & Declarations', 'Certificates & Declarations image.png'),
  mk('Product Information Files',   'Product information files image.png'),
];

console.log('\nMatcher, against the real SharePoint names');
is('the main image is recognised as the lead',
   A.mainPortalImage(IMAGES), 'URL::Main Product portal image');
yes('the main image is NOT treated as a section image',
   !A._portalIsMain(IMAGES[1]) && !A._portalIsMain(IMAGES[2]) && !A._portalIsMain(IMAGES[3]));

const SECS = HUB.productPortal.sections;
const map  = A.assignPortalImages(IMAGES, SECS);
is('Datasheets & MSDS gets its own image',        map.get('data'),  'URL::Datasheets & MSDS');
is('Certificates & Declarations gets its own',    map.get('certs'), 'URL::Certificates & Declarations');
is('Product Information Files gets its own',      map.get('pif'),   'URL::Product Information Files');
is('only three cards are placed',                 map.size, 3);
yes('the lead image never lands on a card',
    ![...map.values()].includes('URL::Main Product portal image'));
yes('sections with no artwork get nothing rather than the wrong thing',
    !map.has('pcn') && !map.has('training') && !map.has('manuals'));

// One picture, one card — the failure that bit the landing images twice.
console.log('\nMatcher, the traps');
const dupes = [...map.values()];
is('no image is used twice', new Set(dupes).size, dupes.length);

// A near-miss must not steal a card: "Certificates" alone should reach
// certs, but must not outrank the exact folder if both are present.
const near = mk('Certificates', 'certificates.png');
const map2 = A.assignPortalImages(IMAGES.concat([near]), SECS);
is('the exact folder still wins over a near-miss',
   map2.get('certs'), 'URL::Certificates & Declarations');

// Marketing renaming a folder must not silently blank a card: a folder
// the matcher cannot place leaves the section empty AND warns.
const junk = mk('Holiday photos', 'IMG_0042.png');
const map3 = A.assignPortalImages([junk], SECS);
is('an unrelated image places nothing', map3.size, 0);

// Sub-type chips: two or more, else none.
console.log('\nOption A — the sub-type chips');
const state = { files: [
  ...Array.from({ length: 4 }, () => ({ _cat: 'Declarations of Conformity' })),
  ...Array.from({ length: 5 }, () => ({ _cat: 'Kitemark certificates' })),
  ...Array.from({ length: 3 }, () => ({ _cat: 'MED' })),
] };
const certs = SECS.find(s => s.key === 'certs');
const chips = A._ppSubTypes({ sec: certs, cat: ['Declarations of Conformity', 'Kitemark certificates', 'MED'] }, state);
is('three types give three chips', chips.length, 3);
is('the biggest type leads', chips[0].label, 'Kitemark certificates');
// The chip carries its position in the band's cat list, not its label.
is('a chip carries its cat index, and it is right',
   ['Declarations of Conformity', 'Kitemark certificates', 'MED'][chips[0].i], 'Kitemark certificates');
yes('every chip index is a number', chips.every(c => typeof c.i === 'number'));
is('a single-type section gets no chips',
   A._ppSubTypes({ sec: certs, cat: ['MED'] }, state).length, 0);
is('a folder-driven section gets no chips',
   A._ppSubTypes({ sec: SECS.find(s => s.key === 'samples'), cat: 'Sample Requests' }, state).length, 0);

console.log(`\n${pass} passed, ${fail} failed\n`);
process.exit(fail ? 1 : 0);
