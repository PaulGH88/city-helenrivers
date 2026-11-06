// Capture the responses the City's own scripts fetch from their server at run time (sub-menus,
// mega-menu content, feature toggles, the anti-forgery token), from the live homepage in the
// visible port-9222 Chrome (never a hidden browser on the G6). Writes snapshot/cp/<name>.
// The copy serves them as files through banner/cp-shim.js, so no City script fails and pops an
// alert box. Run before tools/build.py:   node tools/capture-endpoints.mjs
import fs from 'node:fs';
import path from 'node:path';
const here = path.dirname(new URL(import.meta.url).pathname).replace(/^\/([A-Z]:)/, '$1');
const outDir = path.join(here, '..', 'snapshot', 'cp');
fs.mkdirSync(outDir, { recursive: true });
const base = 'http://127.0.0.1:9222';
const WANT = [
  ['toggle.json', /\/Toggle$/i],
  ['antiforgery.json', /\/antiforgery$/i],
  ['submenus.html', /\/Pages\/MenuMain\/HiddenMainSubMenus/i],
  ['content-load.html', /\/Content\/Load\?/i],
];
const t = await (await fetch(base + '/json/new?about:blank', { method: 'PUT' })).json();
const s = new WebSocket(t.webSocketDebuggerUrl); await new Promise(r => s.onopen = r);
let n = 0; const p = new Map(); const reqs = {}; const done = {};
s.onmessage = m => {
  const d = JSON.parse(m.data);
  if (d.method === 'Network.requestWillBeSent') reqs[d.params.requestId] = d.params.request;
  if (d.method === 'Network.loadingFinished') {
    const r = reqs[d.params.requestId]; if (!r) return;
    const u = r.url.split('&_=')[0];
    const w = WANT.find(([, re]) => re.test(u)); if (!w || done[w[0]]) return;
    done[w[0]] = 'pending';
    send('Network.getResponseBody', { requestId: d.params.requestId }).then(b => {
      const body = b.result.base64Encoded ? Buffer.from(b.result.body, 'base64').toString('utf8') : b.result.body;
      fs.writeFileSync(path.join(outDir, w[0]), body);
      done[w[0]] = body.length;
      console.log('captured', w[0], body.length, 'chars from', r.method, u.slice(0, 120));
    });
  }
  if (d.method === 'Page.javascriptDialogOpening') send('Page.handleJavaScriptDialog', { accept: true });
  if (p.has(d.id)) { p.get(d.id)(d); p.delete(d.id); }
};
const send = (method, params = {}) => new Promise(r => { const i = ++n; p.set(i, r); s.send(JSON.stringify({ id: i, method, params })); });
await send('Network.enable', { maxResourceBufferSize: 20000000, maxTotalBufferSize: 100000000 });
await send('Page.enable');
await send('Page.navigate', { url: 'https://www.cityofdelafieldwi.gov/' });
for (let i = 0; i < 20 && WANT.some(([k]) => typeof done[k] !== 'number'); i++) await new Promise(r => setTimeout(r, 1000));
// The mega-menu content is asked for only when the City's scripts decide to; ask for it the same
// way they do (same URL, same body), from inside the live page.
if (typeof done['content-load.html'] !== 'number') {
  const r = await send('Runtime.evaluate', { awaitPromise: true, returnByValue: true, expression: `(async () => {
    const ids = $('[data-displaymegamenu="True"]').map(function () { return $('#mainNavMegaMenu' + $(this).data('pageid') + ' div.pageContent').data('containerid'); }).toArray();
    const url = '/Content/Load?contentCollectionID=' + $('#hdnContentCollectionID').val() + '&StructureID=' + $('#hdnStructureID').val() + '&themeID=' + $('#hdnThemeID').val();
    const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contentContainerIDs: ids }) });
    return { status: res.status, ids: ids.length, body: await res.text() };
  })()` });
  const v = r.result?.result?.value;
  if (v && v.status === 200 && v.ids > 0) {
    fs.writeFileSync(path.join(outDir, 'content-load.html'), v.body);
    done['content-load.html'] = v.body.length;
    console.log('captured content-load.html', v.body.length, 'chars for', v.ids, 'mega-menu containers');
  } else console.error('content-load failed', v && v.status, v && v.ids);
}
if (typeof done['toggle.json'] !== 'number') {
  const r = await send('Runtime.evaluate', { awaitPromise: true, returnByValue: true, expression: `fetch('/Toggle').then(r => r.text())` });
  const v = r.result?.result?.value;
  if (v && v.startsWith('{')) { fs.writeFileSync(path.join(outDir, 'toggle.json'), v); done['toggle.json'] = v.length; console.log('captured toggle.json', v.length); }
}
await fetch(base + '/json/close/' + t.id);
const missing = WANT.filter(([k]) => typeof done[k] !== 'number').map(([k]) => k);
if (missing.length) { console.error('MISSING', missing.join(', ')); process.exit(1); }
process.exit(0);
