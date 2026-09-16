// Harness for the 16 Sep 2026 second batch — the saved index, and the
// count rail coming back off.
//
//   node "UPLOAD TO GITHUB - 16 Sep 2026 (portal speed 2 + rail)\verify.js"
//
// The saved index is the risky part: it is the first thing in this hub
// that shows the reader something that did not come from Graph a moment
// earlier. So the cache is EXTRACTED and run against a fake
// localStorage — round-tripped, expired, overflowed — rather than read
// and hoped about. (fix-5 rule: sources are decommented first.)

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const DIR = __dirname;
const PREV = path.join(DIR, '..', 'UPLOAD TO GITHUB - 16 Sep 2026 (deck 8 - portal images + option A)');
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
const HTML = fs.readFileSync(path.join(DIR, 'index.html'), 'utf8');
const HTMLC = decomment(HTML.replace(/<!--[\s\S]*?-->/g, ''));
const CFG  = fs.readFileSync(path.join(PREV, 'config.js'), 'utf8');   // unchanged this batch

console.log('\nStatic');
yes('index.html asks for graph.js v24', /graph\.js\?v=24/.test(HTML));
yes('config.js is untouched at v20', /config\.js\?v=20/.test(HTML));

console.log('\nStatic — the rail is off the front');
yes('no rail markup left in the renderer', !/class="px-rail pp-rail"/.test(CODE));
yes('no All-chip left on the front', !/ppOpenSection\(-1\)">All/.test(CODE));
yes('the cards are still the way in', /class="px-grid"/.test(CODE));
yes('the lead still carries the totals', /pp-lead-codes/.test(CODE));
yes('"Search everything" survives in the lead', /cfg\.cta \|\| 'Search everything'/.test(CODE));

console.log('\nStatic — the saved index');
yes('the cache is versioned', /const LIB_CACHE_VERSION\s*=\s*\d+/.test(CODE));
yes('it has a TTL', /const LIB_CACHE_TTL\s*=/.test(CODE));
yes('the crawl is its own function now', /async function _libCrawlAll\(key\)/.test(CODE));
yes('the background refresh exists', /async function _libRefresh\(key\)/.test(CODE));
yes('the refresh is NOT awaited', /_libRefresh\(key\);\s*\/\//.test(RAW));
yes('decoration is shared by both paths', /function _libDecorate\(key, rows\)/.test(CODE));
yes('the crawl no longer writes LIB\\[key\\].files directly',
    !/LIB\[key\]\.files = merged/.test(CODE));
yes('Graph calls are counted', /GRAPH_CALLS\+\+/.test(CODE));
yes('the count is printed with the timing', /graphCallsSince\(c0\)/.test(CODE));

console.log('\nStatic — what must NOT have been lost');
yes('16 Sep portal images still present', /function fetchPortalImages\(\)/.test(CODE));
yes('option A cards still present', /function _ppCard\(l, i, state\)/.test(CODE));
yes('sub-type chips still index-based', /subIndex\s*<\s*bandCats\.length/.test(CODE));
yes('11 Sep links fix still present', /_ppLinksHtml\('\*'\)/.test(CODE));
yes('14 Sep gate still present', /const\s+GRAPH_MAX_INFLIGHT\s*=\s*\d+/.test(CODE));
yes('14 Sep parallel roots still present', /Promise\.all\(src\.roots\.map/.test(CODE));

// ── run the cache for real ─────────────────────────────────────
const cfgBox = { console: { info() {}, warn() {}, log() {} } };
vm.createContext(cfgBox);
vm.runInContext(CFG + '\n;this.__cfg = HUB_CONFIG;', cfgBox);
const HUB = cfgBox.__cfg;

function grab(name) {
  const i = CODE.indexOf('function ' + name + '(');
  if (i < 0) throw new Error('could not find ' + name);
  let d = 0;
  for (let k = CODE.indexOf('{', i); k < CODE.length; k++) {
    if (CODE[k] === '{') d++;
    else if (CODE[k] === '}') { d--; if (!d) return CODE.slice(i, k + 1); }
  }
  throw new Error('unbalanced ' + name);
}

// A fake localStorage with a settable ceiling, so the quota path is a
// test rather than a hope.
function fakeStore(limitBytes) {
  const m = new Map();
  return {
    limit: limitBytes || Infinity,
    getItem: k => (m.has(k) ? m.get(k) : null),
    removeItem: k => m.delete(k),
    setItem(k, v) {
      if (String(v).length > this.limit) {
        const e = new Error('QuotaExceededError'); e.name = 'QuotaExceededError'; throw e;
      }
      m.set(k, String(v));
    },
    _size: () => [...m.values()].reduce((a, b) => a + b.length, 0),
    _keys: () => [...m.keys()],
  };
}

const store = fakeStore();
const box = {
  console: { info() {}, warn() {}, log() {} },
  HUB_CONFIG: HUB, Math, JSON, Date, Set, Map, Array, String, Object, RegExp,
  localStorage: store,
};
vm.createContext(box);
const names = ['_libCfg', '_libSlim', '_libCacheRead', '_libCacheWrite',
               '_libDecorate', '_libTag', '_libCatLabel'];
vm.runInContext(
  CODE.match(/const LIB_CACHE_VERSION[\s\S]*?const _libCacheKey = key => `[^`]+`;/)[0] + '\n' +
  names.map(grab).join('\n') +
  '\n;this.__api = {' + names.join(',') + ', _libCacheKey};', box);
const A = box.__api;

// Rows shaped like the real crawl's output, including the thing that
// must never be written to disk.
const row = (name, p, src) => ({
  id: 'ITEM' + name, _driveId: 'DRIVE1', name, size: 1234,
  lastModifiedDateTime: '2026-09-15T10:00:00Z',
  webUrl: 'https://checkfireltd.sharepoint.com/x/' + name,
  file: { mimeType: 'application/pdf', hashes: { quickXorHash: 'x' } },
  '@microsoft.graph.downloadUrl': 'https://SHORT-LIVED/' + name,
  _path: p, _source: src, _sourceKey: 'k',
});

const LIVE = A._libDecorate('product', [
  row('Commander 6kg Powder Declaration of Conformity.pdf', ['DOCs'], 'Product Portal'),
  row('CommanderEDGE 3ltr Wet Chem Data Sheet.pdf', ['Product Documents'], 'Sales & Marketing'),
  row('Kitemark Certificate CO2 2kg.pdf', ['Kitemark Certificates'], 'Product Portal'),
  row('Commander PIF.pdf', ['01. Marketing', '08. PDF PIF, Data Sheets, MSDS Sheets & Toolkits'], 'Sales & Marketing'),
]);

console.log('\nThe saved index, round-tripped');
A._libCacheWrite('product', LIVE);
const back = A._libCacheRead('product');
yes('something was written', !!back);
is('every row came back', back.files.length, LIVE.length);

const RESTORED = A._libDecorate('product', back.files);
is('the categories survive the round trip',
   RESTORED.map(f => f._cat).join('|'), LIVE.map(f => f._cat).join('|'));
is('the tags survive the round trip',
   RESTORED.map(f => f._tag).join('|'), LIVE.map(f => f._tag).join('|'));
is('the PIF is still a PIF and not an MSDS', RESTORED[3]._cat, 'PIF');

console.log('\nThe saved index, the traps');
const raw = store.getItem(A._libCacheKey('product'));
yes('the short-lived download URL is NOT on disk', !/SHORT-LIVED/.test(raw));
yes('_driveId IS on disk — a download needs it to re-fetch the URL', /DRIVE1/.test(raw));
yes('item ids are on disk', /ITEM/.test(raw));
yes('nothing is stored under the wrong key',
    store._keys().every(k => k === A._libCacheKey('product')));

// Expiry: hand it an entry older than the TTL.
const TTL = Number((CODE.match(/const LIB_CACHE_TTL\s*=\s*([\d\s*]+);/) || [])[1]
  .split('*').map(s => Number(s.trim())).reduce((a, b) => a * b, 1));
store.setItem(A._libCacheKey('product'),
  JSON.stringify({ t: Date.now() - TTL - 1000, v: LIVE.map(A._libSlim) }));
is('a stale index is ignored', A._libCacheRead('product'), null);

// Corrupt / empty entries must not throw.
store.setItem(A._libCacheKey('product'), 'not json at all');
is('a corrupt index is ignored, not thrown', A._libCacheRead('product'), null);
store.setItem(A._libCacheKey('product'), JSON.stringify({ t: Date.now(), v: [] }));
is('an empty index is ignored', A._libCacheRead('product'), null);

// Quota: writing must fail quietly AND clear up after itself, because a
// half-written entry read back as the whole library is the worst
// outcome available here.
const tiny = fakeStore(50);
box.localStorage = tiny;
vm.runInContext('localStorage = this.localStorage;', box);
let threw = false;
try { A._libCacheWrite('product', LIVE); } catch (_) { threw = true; }
yes('a full localStorage does not throw', !threw);
is('and leaves nothing behind', tiny._keys().length, 0);

// Size: the real portal is ~940 files. It has to fit with room to spare.
box.localStorage = store;
vm.runInContext('localStorage = this.localStorage;', box);
const many = [];
for (let i = 0; i < 940; i++) {
  many.push(row(`Commander ${i} Declaration of Conformity.pdf`, ['DOCs', 'CO2'], 'Product Portal'));
}
A._libCacheWrite('product', A._libDecorate('product', many));
const bytes = store.getItem(A._libCacheKey('product')).length;
yes(`940 files serialise to ${(bytes / 1024).toFixed(0)} KB — well inside a 5 MB quota`,
    bytes < 1024 * 1024, `${bytes} bytes`);
is('and all 940 read back', A._libCacheRead('product').files.length, 940);

console.log(`\n${pass} passed, ${fail} failed\n`);
process.exit(fail ? 1 : 0);
