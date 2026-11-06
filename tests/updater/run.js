// Updater test bed, step 1: run every test workbook through the updater with the candidate library.
// Usage: TESTBED_DIRS='name=/folder;...' [GEO_CONFIG='{...}'] node tests/updater/run.js [library.txt]      (needs: npm i playwright; a Chromium build)
// Writes tests/updater/out/<name> (updated copies, not committed) and tests/updater/out/run.json.
const fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const ROOT = path.resolve(__dirname, '../..');
// LIB: library text to install (default: the published geoXL snapshot); GEO_CONFIG: the page's window.GeoConfig as JSON.
const LIB = process.argv[2] || path.join(ROOT, 'docs/library/releases', JSON.parse(fs.readFileSync(path.join(ROOT,'docs/library/current.json'),'utf8')).version, 'library.txt');
const CONFIG = process.env.GEO_CONFIG || '{}';
const OUT = process.env.TESTBED_OUT || path.join(__dirname, 'out');
// TESTBED_DIRS="name=/path/to/folder;name2=/other" checks other folders (e.g. the Drive master folder)
const sets = process.env.TESTBED_DIRS ? process.env.TESTBED_DIRS.split(';').map(x => x.split('=')) :
  [];   // set TESTBED_DIRS; workbooks are kept outside this public repository
const exe = process.env.CHROMIUM || undefined;
(async () => {
  fs.rmSync(OUT, { recursive: true, force: true }); fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch(exe ? { executablePath: exe } : {});
  const server = require('http').createServer((q, res) => { res.setHeader('content-type', 'text/html'); res.end('<!doctype html><title>t</title>'); });
  await new Promise(r => server.listen(0, '127.0.0.1', r));
  const page = await browser.newPage();
  await page.goto('http://localhost:' + server.address().port + '/');   // secure context (crypto.subtle)
  await page.addScriptTag({ content: fs.readFileSync(path.join(ROOT, 'tools/vendor/jszip.min.js'), 'utf8') });
  await page.addScriptTag({ content: 'window.GeoConfig=' + CONFIG + ';' });
  await page.addScriptTag({ content: fs.readFileSync(path.join(ROOT, 'tools/updater-core.js'), 'utf8') });
  const libText = fs.readFileSync(LIB, 'utf8');
  const libInfo = await page.evaluate(async t => { window.L = await GeoUpdater.readLibrary(t); return { version: L.version, hash: L.hash, n: L.all.length,
    defs: L.all.map(d => ({ name: d.name, lambda: d.ast.k === 'call' && d.ast.v.toUpperCase() === 'LAMBDA', compiled: d.compiled })) }; }, libText);
  const results = [];
  for (const [set, dir] of sets) for (const f of fs.readdirSync(dir).filter(f => /\.xlsx$/i.test(f) && !f.startsWith('~$')).sort()) {
    const src = path.join(dir, f), b64 = fs.readFileSync(src).toString('base64');
    const r = await page.evaluate(async ([b64, name]) => {
      const run = async bytes => { const r = await GeoUpdater.updateWorkbook(bytes, name, L); return r; };
      const toB64 = u8 => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); };
      try {
        const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        const first = await run(bytes.buffer);
        const second = await run(first.bytes.buffer);           // idempotency: update the updated copy again
        return { ok: true, report: first.report, second: second.report.counts, out: toB64(first.bytes), out2: toB64(second.bytes) };
      } catch (e) { return { ok: false, error: String(e.message || e) }; }
    }, [b64, f]);
    if (r.ok) { fs.writeFileSync(path.join(OUT, f), Buffer.from(r.out, 'base64')); fs.writeFileSync(path.join(OUT, f.replace(/\.xlsx$/i, '.second.xlsx')), Buffer.from(r.out2, 'base64')); delete r.out; delete r.out2; }
    results.push({ set, file: f, source: src, ...r });
    console.log((r.ok ? 'updated ' : 'FAILED  ') + set + '/' + f + (r.ok ? '' : ': ' + r.error));
  }
  fs.writeFileSync(path.join(OUT, 'run.json'), JSON.stringify({ library: libInfo, results }, null, 1));
  await browser.close(); server.close();
})();
