/* Browser QA with isolated fixtures. Never writes to the live Apps Script service. */
const fs = require('node:fs');
const path = require('node:path');
const { spawn } = require('node:child_process');
const assert = require('node:assert/strict');
fs.mkdirSync('.qa',{recursive:true});
const log=(message)=>{fs.appendFileSync('.qa/browser.log',message+'\n');console.log(message);};
log('Starting browser QA');
process.on('exit',(code)=>fs.appendFileSync('.qa/browser.log','Exit '+code+'\n'));
const chromePath = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const preview = spawn(process.execPath,['scripts/preview.cjs'],{windowsHide:true,stdio:'ignore'});
const browser = spawn(chromePath,['--headless=new','--disable-gpu','--no-first-run','--no-default-browser-check','--disable-extensions','--remote-debugging-port=9224',`--user-data-dir=${path.resolve('.qa/chrome')}`,'about:blank'],{windowsHide:true,stdio:'ignore'});
const delay = (ms) => new Promise((r) => setTimeout(r,ms));
let ws;
async function connect() {
  for (let i=0;i<60;i++) {
    try { const pages = await (await fetch('http://127.0.0.1:9224/json')).json(); if (pages[0]) return pages[0].webSocketDebuggerUrl; } catch {}
    await delay(200);
  }
  throw new Error('Chrome did not start');
}
const fixture = `(() => {
  const rifa={rifaId:'boleta',titulo:'Rifa de $500.000',premio:'$500.000',cantidad:100,precio:20000,loteria:'Boyacá',fecha:'2026-10-31',hora:'17:00',responsable:'Arelys B. Mendoza B.'};
  const records=[{numeroId:'00',estado:'apartado',clienteNombre:'Ana Pérez',clienteTelefono:'3001234567',revision:1,updatedAt:'2026-10-20T21:01:02Z'}, {numeroId:'01',estado:'pagado',clienteNombre:'Carlos',clienteTelefono:'3001234567',paymentMethod:'transferencia',revision:1}, {numeroId:'02',estado:'pagado',clienteNombre:'Luisa',clienteTelefono:'3001234567',paymentMethod:'efectivo',revision:1}];
  const pub=(n)=>({numeroId:n.numeroId,estado:n.estado==='libre'?'libre':'apartado',clienteNombre:n.clienteNombre,updatedAt:n.updatedAt});
  const append=Element.prototype.appendChild;
  Element.prototype.appendChild=function(node){
    if(node.tagName==='SCRIPT' && node.src.startsWith('https://script.google.com')){
      const u=new URL(node.src), action=u.searchParams.get('action'), cb=u.searchParams.get('callback');
      queueMicrotask(()=>window[cb]?.({ok:true,protocol:2,data:action==='getRifas'?[rifa]:records.map(pub)})); return node;
    }
    return append.call(this,node);
  };
  const realFetch=window.fetch;
  window.fetch=async function(url,options){
    if(!String(url).startsWith('https://script.google.com')) return realFetch.call(this,url,options);
    if(!navigator.onLine) throw new TypeError('offline');
    const {action,payload:p}=JSON.parse(options.body); let result={ok:true};
    if(action==='login')result=p.pin==='4826'?{ok:true,token:'qa-token',expiresAt:Date.now()+1800000}:{ok:false,error:'Contraseña incorrecta.'};
    if(action==='adminSnapshot'){
      const paid=records.filter(n=>n.estado==='pagado'),apartados=records.filter(n=>n.estado==='apartado').length;
      result={ok:true,data:records,summary:{libres:100-paid.length-apartados,apartados,pagados:paid.length,recaudado:paid.length*20000,meta:2000000,transferencia:paid.filter(n=>n.paymentMethod==='transferencia').length*20000,efectivo:paid.filter(n=>n.paymentMethod==='efectivo').length*20000,sinClasificar:0}};
    }
    if(action==='reserveNumber'||action==='updateNumber'){
      if(p.numeroId==='08')result={ok:false,code:'CONFLICT',record:{numeroId:'08',clienteNombre:'Marta',updatedAt:'2026-10-20T21:01:02Z'}};
      else{ const existing=records.findIndex(n=>n.numeroId===p.numeroId); const next={...p,estado:action==='reserveNumber'?'apartado':p.estado,revision:1}; if(existing>=0)records[existing]=next;else records.push(next); }
    }
    return new Response(JSON.stringify(result),{headers:{'Content-Type':'application/json'}});
  };
})()`;
(async () => {
  let seq=0; const pending=new Map(); const errors=[];
  try {
    const target=await connect();log('Chrome connected');ws=new WebSocket(target); await new Promise((r)=>ws.addEventListener('open',r,{once:true}));log('CDP ready');
    ws.addEventListener('message',(event)=>{const msg=JSON.parse(event.data);if(msg.id){const p=pending.get(msg.id);pending.delete(msg.id);msg.error?p.reject(new Error(msg.error.message)):p.resolve(msg.result);}if(msg.method==='Runtime.exceptionThrown')errors.push(msg.params.exceptionDetails.text);});
    const send=(method,params={})=>new Promise((resolve,reject)=>{const id=++seq;pending.set(id,{resolve,reject});ws.send(JSON.stringify({id,method,params}));});
    const evaluate=async(expression)=>{const r=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
    const wait=async(expression)=>{for(let i=0;i<100;i++){if(await evaluate(expression))return;await delay(100);}throw new Error('Timed out: '+expression);};
    const click=(id)=>evaluate(`document.querySelector(${JSON.stringify(id)}).click()`);
    const set=(id,value)=>evaluate(`document.querySelector(${JSON.stringify(id)}).value=${JSON.stringify(value)}`);
    const screenshot=async(name)=>{const r=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});fs.writeFileSync('.qa/'+name+'.png',Buffer.from(r.data,'base64'));};
    await send('Page.enable'); await send('Runtime.enable'); await send('Page.addScriptToEvaluateOnNewDocument',{source:fixture});
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1100,deviceScaleFactor:1,mobile:false});
    await send('Page.navigate',{url:'http://127.0.0.1:4173/RIFA-PRO.COM/'});
    await wait('document.querySelectorAll(".numero").length===100');log('Page ready');
    assert(await evaluate('document.getElementById("adminSummary").hidden'));
    assert.equal(await evaluate('document.getElementById("kpiRecaudado").textContent'),'');
    assert(await evaluate('document.documentElement.scrollWidth<=innerWidth'));
    await screenshot('desktop-public');
    await click('[data-number="07"]'); await set('#clienteNombre','Pedro QA'); await set('#clienteTelefono','3001234567'); await click('#saveNumberButton');
    await wait('document.getElementById("noticeDialog").open'); assert((await evaluate('document.getElementById("noticeTitle").textContent')).includes('es tuyo'));
    await click('[data-close="noticeDialog"]'); await wait(`document.querySelector('[data-number="07"]').classList.contains('apartado')`);
    await click('[data-number="08"]'); await set('#clienteNombre','Pedro QA'); await set('#clienteTelefono','3001234567'); await click('#saveNumberButton');
    await wait('document.getElementById("noticeDialog").open'); assert((await evaluate('document.getElementById("noticeMessage").textContent')).includes('Marta'));
    await click('[data-close="noticeDialog"]');
    await click('#adminButton'); await set('#pinInput','9999'); await click('#loginButton'); await wait('document.getElementById("adminError").textContent.includes("incorrecta")');
    assert(await evaluate('document.getElementById("adminSummary").hidden')); await set('#pinInput','4826'); await click('#loginButton');
    await wait('!document.getElementById("adminSummary").hidden');
    await click('[data-number="09"]'); await set('#clienteNombre','Diana QA'); await set('#clienteTelefono','3001234567'); await set('#estadoSelect','pagado');
    await evaluate('document.getElementById("estadoSelect").dispatchEvent(new Event("change"))');
    assert(await evaluate('!document.getElementById("paymentField").hidden'));
    assert.equal(await evaluate('document.getElementById("numberForm").checkValidity()'),false);
    await click('input[value="transferencia"]'); assert(await evaluate('document.getElementById("numberForm").checkValidity()'));
    await screenshot('payment-transfer'); await click('#saveNumberButton'); await wait('document.getElementById("noticeDialog").open'); await click('[data-close="noticeDialog"]');
    await wait('document.getElementById("totalTransfer").textContent.includes("40.000")');
    await click('[data-number="10"]'); await set('#clienteNombre','Luis QA'); await set('#clienteTelefono','3001234567'); await set('#estadoSelect','pagado'); await evaluate('document.getElementById("estadoSelect").dispatchEvent(new Event("change"))'); await click('input[value="efectivo"]'); await click('#saveNumberButton');
    await wait('document.getElementById("noticeDialog").open'); await click('[data-close="noticeDialog"]'); await wait('document.getElementById("totalCash").textContent.includes("40.000")');
    await screenshot('desktop-admin');
    assert(!(await evaluate('localStorage.getItem("rifapro.public.v2")')).includes('3001234567'));
    await click('#adminButton'); await wait('document.getElementById("adminSummary").hidden');
    assert.equal(await evaluate('document.getElementById("kpiRecaudado").textContent'),'');
    for(const width of [320,375,390,768,1024]){
      await send('Emulation.setDeviceMetricsOverride',{width,height:844,deviceScaleFactor:1,mobile:width<768}); await delay(150);
      assert(await evaluate('document.documentElement.scrollWidth<=innerWidth'),`horizontal overflow at ${width}`);
      if(width===390) await screenshot('mobile-public');
    }
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true}); await click('#themeToggle'); await screenshot('mobile-dark');
    await click('#themeToggle');
    await wait('navigator.serviceWorker.controller !== null');
    const assets=await evaluate(`Promise.all(['./images/icon-192.png','./images/icon-512.png','./images/apple-touch-icon.png'].map(src=>new Promise(resolve=>{const i=new Image();i.onload=()=>resolve(i.naturalWidth);i.onerror=()=>resolve(0);i.src=src})))`);
    assert.deepEqual(assets,[192,512,180]);
    await send('Network.enable'); await send('Network.emulateNetworkConditions',{offline:true,latency:0,downloadThroughput:0,uploadThroughput:0});
    await wait('!navigator.onLine'); await click('[data-number="11"]'); await wait('document.getElementById("noticeDialog").open');
    assert((await evaluate('document.getElementById("noticeTitle").textContent')).includes('Conéctate'));
    await send('Page.reload'); await wait('document.getElementById("dashboard")!==null');
    assert.equal(await evaluate('document.title'),'Rifa Pro · Tu próximo número de la suerte');
    assert.equal(errors.length,0,JSON.stringify(errors));
    log('PASS: public/admin isolation, confirmation, conflict modal, PIN, required payment method, transfer/cash totals, logout scrubbing, 320/375/390/768/1024 layouts, PWA icons and offline shell.');
    console.log('Screenshots saved in .qa/. Fixture backend only; live deployment still needs validation.');
  } finally { if(ws?.readyState===1)ws.close();browser.kill();preview.kill(); }
})().catch((e)=>{log(e.stack);process.exitCode=1;});
