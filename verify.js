// Harness for the 14 Sep 2026 portal-speed batch.
//
//   node "UPLOAD TO GITHUB - 14 Sep 2026 (portal speed)\verify.js"
//
// It does not read the plumbing and hope. It EXTRACTS the shipped Graph
// layer out of graph.js, runs it against a fake Graph, and measures what
// it actually did: peak concurrency, retries, cache hits, request count.
// (fix-5 rule: every source is decommented before any assertion, so a
// comment can never satisfy one.)

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
const HTML = fs.readFileSync(path.join(DIR, 'index.html'), 'utf8');

// ── static assertions ──────────────────────────────────────────
console.log('\nStatic');
yes('index.html asks for graph.js v22', /graph\.js\?v=22/.test(HTML));
yes('gate constant is declared', /const\s+GRAPH_MAX_INFLIGHT\s*=\s*\d+/.test(CODE));
yes('roots run in parallel', /Promise\.all\(src\.roots\.map/.test(CODE));
yes('drives run in parallel', /Promise\.all\(drives\.map/.test(CODE));
yes('no sequential root loop left', !/for\s*\(const root of src\.roots\)/.test(CODE));
yes('no sequential drive loop left', !/for\s*\(const drive of drives\)/.test(CODE));
yes('crawl no longer swallows errors', !/catch\s*\(_\)\s*\{\s*return;\s*\}/.test(CODE));
yes('FormServerTemplates is NOT skipped', !/formservertemplates/i.test(
  (CODE.match(/const SKIP_LIBRARIES\s*=\s*\[[\s\S]*?\]/) || [''])[0]));
yes('Preservation Hold IS skipped', /preservation hold library/.test(
  (CODE.match(/const SKIP_LIBRARIES\s*=\s*\[[\s\S]*?\]/) || [''])[0]));
yes('library load is timed', /Date\.now\(\)\s*-\s*t0/.test(CODE));
// Baseline check — the fix-4/fix-5 trap was building on the wrong file.
// 11 Sep's links fix must still be in here.
yes("11 Sep links fix is still present", /_ppLinksHtml\('\*'\)/.test(CODE));
yes("11 Sep front-mode is still present", /const\s+front\s*=\s*sectionKey\s*===\s*'\*'/.test(CODE));

// ── extract the Graph layer and run it ─────────────────────────
const start = CODE.indexOf("const GRAPH_BASE");
const endM  = CODE.indexOf('async function _fetchDriveChildrenLive');
const endTail = CODE.indexOf('\n}', CODE.indexOf('return (data.value || []).map', endM));
if (start < 0 || endM < 0 || endTail < 0) {
  bad('could locate the Graph layer in graph.js');
  process.exit(1);
}
const slice = CODE.slice(start, endTail + 2);

// Fake Graph. Records every call, how many were in flight at once, and
// serves a scripted 429 for one path so the retry path is exercised.
const log = [];
let inflight = 0, peak = 0, seen429 = 0;
const throttleOnce = new Set(['/drives/D/items/THROTTLE/children']);

async function fakeFetch(url) {
  const p = url.replace('https://graph.microsoft.com/v1.0', '').split('?')[0];
  log.push(p);
  inflight++; peak = Math.max(peak, inflight);
  await new Promise(r => setTimeout(r, 15));
  inflight--;
  if (throttleOnce.has(p)) {
    throttleOnce.delete(p);
    seen429++;
    return { ok: false, status: 429, headers: { get: () => '0.02' }, json: async () => ({}) };
  }
  return {
    ok: true, status: 200, headers: { get: () => null },
    json: async () => ({ value: [{ id: 'x', name: 'a.pdf', file: {} }] }),
  };
}

const sandbox = {
  console: { info() {}, warn() {}, log() {} },
  setTimeout, clearTimeout, URL, Date, Math, JSON, Promise, parseFloat, isFinite,
  fetch: fakeFetch,
  getAccessToken: async () => 'tok',
  HUB_CONFIG: { sharepointSite: 'https://x.sharepoint.com/sites/a' },
  sessionStorage: (() => {
    const m = new Map();
    return { getItem: k => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, v) };
  })(),
};
vm.createContext(sandbox);
vm.runInContext(slice + '\n;this.__api = { graphFetch, fetchDriveChildren, GRAPH_MAX_INFLIGHT };', sandbox);
const API = sandbox.__api;

(async () => {
  console.log('\nGate and cache, measured');

  // 40 distinct folders at once — the shape of the real crawl now that
  // roots and drives run together.
  const many = Array.from({ length: 40 }, (_, i) => `F${i}`);
  await Promise.all(many.map(id => API.fetchDriveChildren('D', id)));
  yes(`peak concurrency ${peak} never exceeded the gate (${API.GRAPH_MAX_INFLIGHT})`,
      peak <= API.GRAPH_MAX_INFLIGHT, `peak was ${peak}`);
  is('40 folders cost 40 requests', log.length, 40);

  // Same folders again — every one should now be served from cache.
  const before = log.length;
  await Promise.all(many.map(id => API.fetchDriveChildren('D', id)));
  is('a second pass costs 0 requests', log.length - before, 0);

  // Ten callers racing for one uncached folder share a single request.
  const b2 = log.length;
  await Promise.all(Array.from({ length: 10 }, () => API.fetchDriveChildren('D', 'RACE')));
  is('10 concurrent callers cost 1 request', log.length - b2, 1);

  // A 429 is retried rather than surfacing as an empty folder.
  const rows = await API.fetchDriveChildren('D', 'THROTTLE');
  is('the 429 was actually served', seen429, 1);
  yes('a throttled folder still returns its files', Array.isArray(rows) && rows.length === 1);

  // root vs a folder called "root" must not collide.
  const b3 = log.length;
  await API.fetchDriveChildren('D2', null);
  await API.fetchDriveChildren('D2', 'root');
  is('null root and a folder named root are separate keys', log.length - b3, 2);

  console.log(`\n${pass} passed, ${fail} failed\n`);
  process.exit(fail ? 1 : 0);
})();
