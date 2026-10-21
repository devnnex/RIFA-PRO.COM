/* Real service-worker upgrade against a local server and isolated raffle fixtures. */
const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const root = path.resolve(__dirname, '..');
const fixture = '(' + require('../tests/browser-fixture.cjs').bootstrap.toString() + ')()';
const types = { '.html':'text/html', '.js':'text/javascript', '.css':'text/css', '.json':'application/json', '.png':'image/png', '.svg':'image/svg+xml' };
let release = 'qa-v1', ws, browser;
const server = http.createServer((req, res) => {
  const relative = new URL(req.url, 'http://localhost').pathname.replace(/^\/RIFA-PRO\.COM\//, '') || 'index.html';
  if (!/^(index\.html|[a-z-]+\.(js|css|json)|images\/[a-z0-9-]+\.(svg|png))$/.test(relative)) { res.writeHead(404); return res.end(); }
  fs.readFile(path.join(root, relative), (error, bytes) => {
    if (error) { res.writeHead(404); return res.end(); }
    let body = bytes;
    if (relative === 'sw.js') body = bytes.toString().replace(/rifapro-shell-v[\d.]+/, 'rifapro-shell-' + release);
    if (relative === 'experience.js') body = bytes.toString() + '\nwindow.__release=' + JSON.stringify(release) + ';';
    res.writeHead(200, { 'Content-Type':(types[path.extname(relative)] || 'text/plain') + (relative.endsWith('.png') ? '' : '; charset=utf-8'), 'Cache-Control':'no-store' }); res.end(body);
  });
});
(async () => {
  await new Promise(resolve => server.listen(4181, '127.0.0.1', resolve));
  browser = spawn('C:/Program Files/Google/Chrome/Application/chrome.exe', ['--headless=new','--disable-gpu','--no-first-run','--disable-extensions','--remote-debugging-port=9227', '--user-data-dir=' + path.join(root, '.qa', 'pwa-upgrade-' + Date.now()), 'about:blank'], { windowsHide:true, stdio:'ignore' });
  try {
    let target;
    for (let i = 0; i < 60 && !target; i++) {
      try { target = (await (await fetch('http://127.0.0.1:9227/json')).json()).find(p => p.type === 'page')?.webSocketDebuggerUrl; } catch {}
      if (!target) await delay(200);
    }
    assert(target, 'Chrome starts');
    ws = new WebSocket(target); await new Promise(resolve => ws.addEventListener('open', resolve, { once:true }));
    let sequence = 0; const pending = new Map();
    ws.addEventListener('message', event => { const m = JSON.parse(event.data); if (m.id) { const p = pending.get(m.id); pending.delete(m.id); m.error ? p.reject(new Error(m.error.message)) : p.resolve(m.result); } });
    const send = (method, params = {}) => new Promise((resolve,reject) => { const id = ++sequence; pending.set(id,{resolve,reject}); ws.send(JSON.stringify({id,method,params})); });
    const evaluate = async expression => { const r = await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true}); if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description); return r.result.value; };
    const wait = async expression => { for (let i=0;i<150;i++) { if (await evaluate(expression)) return; await delay(100); } throw new Error('Timed out: ' + expression); };
    await send('Page.enable'); await send('Runtime.enable');
    await send('Page.addScriptToEvaluateOnNewDocument',{source:fixture});
    await send('Page.navigate',{url:'http://127.0.0.1:4181/RIFA-PRO.COM/?rifa=boleta'});
    await wait('window.__release==="qa-v1"&&navigator.serviceWorker.controller!==null&&state.online');
    await send('Page.reload'); await wait('window.__release==="qa-v1"&&state.online');
    await evaluate('setTheme(true);localStorage.setItem("qa.keep","retained");openNumber("12");document.getElementById("clienteNombre").value="Formulario intacto"');
    release = 'qa-v2';
    await evaluate('navigator.serviceWorker.getRegistration().then(r=>r.update()).then(()=>true)');
    await wait('appUpdateReady&&!document.getElementById("appUpdateNotice").hidden');
    assert(await evaluate('window.__release==="qa-v1"&&document.getElementById("numberDialog").open&&document.getElementById("clienteNombre").value==="Formulario intacto"'));
    await evaluate('applyAppUpdate()');
    assert(await evaluate('window.__release==="qa-v1"&&document.getElementById("numberDialog").open'));
    await evaluate('document.getElementById("numberDialog").close();document.getElementById("applyAppUpdate").click()');
    await wait('window.__release==="qa-v2"&&state.online');
    assert(await evaluate('location.search==="?rifa=boleta"&&localStorage.getItem("qa.keep")==="retained"&&document.body.classList.contains("dark-mode")'));
    assert.equal(await evaluate('navigator.serviceWorker.getRegistrations().then(r=>r.length)'),1);
    assert.deepEqual(await evaluate('caches.keys()'),['rifapro-shell-qa-v2']);
    await send('Network.enable');
    await send('Network.emulateNetworkConditions',{offline:true,latency:0,downloadThroughput:0,uploadThroughput:0});
    await send('Page.reload');await wait('window.__release==="qa-v2"&&document.getElementById("dashboard")!==null');
    console.log('PASS: real PWA upgrade v1 → v2; complete shell installed; modal and typed data preserved until update; same installation/link/theme/storage; old cache removed; upgraded shell opens offline.');
  } finally { if (ws?.readyState===1) ws.close(); browser?.kill(); server.close(); }
})().catch(error => { console.error(error); process.exitCode=1; });
