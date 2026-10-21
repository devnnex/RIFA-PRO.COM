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
const browser = spawn(chromePath,['--headless=new','--disable-gpu','--no-first-run','--no-default-browser-check','--disable-extensions','--remote-debugging-port=9224',`--user-data-dir=${path.resolve('.qa/chrome-'+Date.now())}`,'about:blank'],{windowsHide:true,stdio:'ignore'});
const delay = (ms) => new Promise((r) => setTimeout(r,ms));
let ws;
async function connect() {
  for (let i=0;i<60;i++) {
    try { const pages = await (await fetch('http://127.0.0.1:9224/json')).json(); const page=pages.find(p=>p.type==='page'); if (page) return page.webSocketDebuggerUrl; } catch {}
    await delay(200);
  }
  throw new Error('Chrome did not start');
}
const fixture = '(' + require('../tests/browser-fixture.cjs').bootstrap.toString() + ')()';
(async () => {
  let seq=0; const pending=new Map(); const errors=[];
  try {
    const target=await connect();log('Chrome connected');ws=new WebSocket(target); await new Promise((r)=>ws.addEventListener('open',r,{once:true}));log('CDP ready');
    ws.addEventListener('message',(event)=>{const msg=JSON.parse(event.data);if(msg.id){const p=pending.get(msg.id);pending.delete(msg.id);msg.error?p.reject(new Error(msg.error.message)):p.resolve(msg.result);}if(msg.method==='Runtime.exceptionThrown')errors.push(msg.params.exceptionDetails.text);});
    const send=(method,params={})=>new Promise((resolve,reject)=>{const id=++seq;pending.set(id,{resolve,reject});ws.send(JSON.stringify({id,method,params}));});
    const evaluate=async(expression)=>{const r=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(r.exceptionDetails)throw new Error(r.exceptionDetails.exception?.description||r.exceptionDetails.text);return r.result.value;};
    const wait=async(expression)=>{for(let i=0;i<100;i++){if(await evaluate(expression))return;await delay(100);}log('Diagnostic '+JSON.stringify(await evaluate('({url:location.href,ready:document.readyState,title:document.title,config:!!window.RIFA_CONFIG,connection:document.getElementById("connectionText")?.textContent,calls:window.__qa?.calls?.map(c=>c.action),state:typeof state!=="undefined"?{online:state.online,protocol:state.protocol,count:state.rifas.length,active:state.activeId}:null})')));log('Runtime errors '+JSON.stringify(errors));throw new Error('Timed out: '+expression);};
    const click=(id)=>evaluate(`document.querySelector(${JSON.stringify(id)}).click()`);
    const set=(id,value)=>evaluate(`document.querySelector(${JSON.stringify(id)}).value=${JSON.stringify(value)}`);
    const screenshot=async(name)=>{const r=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false});fs.writeFileSync('.qa/'+name+'.png',Buffer.from(r.data,'base64'));};
    await send('Page.enable'); await send('Runtime.enable'); await send('Page.addScriptToEvaluateOnNewDocument',{source:fixture});
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1100,deviceScaleFactor:1,mobile:false});
    await send('Page.navigate',{url:'http://127.0.0.1:4173/RIFA-PRO.COM/?rifa=boleta&slow=1'});
    await wait('typeof state!=="undefined"&&state.entryLoading&&!document.getElementById("entryLoader").hidden');
    assert((await evaluate('document.getElementById("entryLoader").textContent')).includes('Preparando números'));await screenshot('numbers-loader');
    await wait('document.querySelectorAll(".numero").length===100 && state.online');log('Page ready');
    assert(await evaluate('document.getElementById("entryLoader").hidden&&document.getElementById("shareButton").hidden'));
    assert.deepEqual(await evaluate('window.__qa.calls.map(c=>c.action)'),['getSnapshot']);
    assert(await evaluate('document.getElementById("adminSummary").hidden'));
    assert.equal(await evaluate('document.getElementById("kpiRecaudado").textContent'),'');
    assert(await evaluate('document.documentElement.scrollWidth<=innerWidth'));
    await evaluate('window.__stableTile=document.querySelector(".numero");window.__qa.calls=[]');
    await evaluate('refresh()');assert(await evaluate('window.__stableTile===document.querySelector(".numero")'));
    assert.equal(await evaluate('window.__qa.calls.filter(c=>c.action==="getSnapshot").length'),1);
    await screenshot('desktop-public');
    await click('[data-number="07"]');await set('#clienteNombre','Pedro QA');await set('#clienteTelefono','3001234567');
    const instant=await evaluate('(()=>{const start=performance.now();document.getElementById("saveNumberButton").click();return {ms:performance.now()-start,closed:!document.getElementById("numberDialog").open,pending:document.querySelector(\'[data-number="07"]\').classList.contains("pending")}})()');
    assert(instant.closed && instant.pending && instant.ms<100,JSON.stringify(instant));
    await wait('document.getElementById("noticeDialog").open');assert((await evaluate('document.getElementById("noticeTitle").textContent')).includes('es tuyo'));await click('[data-close="noticeDialog"]');
    await click('[data-number="08"]');await set('#clienteNombre','Pedro QA');await set('#clienteTelefono','3001234567');await click('#saveNumberButton');
    await wait('document.getElementById("noticeDialog").open');assert((await evaluate('document.getElementById("noticeMessage").textContent')).includes('Marta'));await click('[data-close="noticeDialog"]');
    await evaluate('openPinLogin(false)');assert(await evaluate('document.getElementById("adminDialog").open'));await set('#pinInput','9999');await click('#loginButton');await wait('document.getElementById("adminError").textContent.includes("incorrecto")');
    assert(await evaluate('document.getElementById("adminSummary").hidden'));await set('#pinInput','4826');await click('#loginButton');await wait('!document.getElementById("adminSummary").hidden');
    const prepare=async(n,method)=>{await click('[data-number="'+n+'"]');await set('#clienteNombre','Cliente '+n);await set('#clienteTelefono','3001234567');await click('[data-state="pagado"]');if(method)await click('input[value="'+method+'"]');};
    await click('#themeToggle');
    await wait(`getComputedStyle(document.querySelector('[data-number="01"]')).backgroundColor==='rgb(67, 38, 83)'`);
    assert.equal(await evaluate(`getComputedStyle(document.querySelector('[data-number="01"]')).backgroundColor`),'rgb(67, 38, 83)');await screenshot('dark-nequi-payments');await click('#themeToggle');
    await prepare('09');assert.equal(await evaluate('document.getElementById("numberForm").checkValidity()'),false);await click('input[value="transferencia"]');
    await screenshot('payment-transfer');await click('#saveNumberButton');assert(await evaluate('!document.getElementById("numberDialog").open'));await wait('state.pending.size===0');
    assert((await evaluate('document.getElementById("totalTransfer").textContent')).includes('40.000'));
    await prepare('10','efectivo');await click('#saveNumberButton');await wait('state.pending.size===0');assert((await evaluate('document.getElementById("totalCash").textContent')).includes('40.000'));
    await click('[data-edit-prize]');await set('#editPremio','$800.000');await evaluate('document.getElementById("editPremio").dispatchEvent(new Event("input"))');
    assert.equal(await evaluate('document.getElementById("prizePreview").textContent'),'$800.000');await set('#editTitulo','Rifa de $800.000');await evaluate('window.__qa.delay=700');
    await evaluate('document.getElementById("prizeForm").requestSubmit()');
    assert.equal(await evaluate('document.querySelector(".ticket-prize").textContent'),'$800.000');assert(await evaluate('!document.getElementById("prizeDialog").open && state.pendingRaffles.size===1'));
    await wait('state.pendingRaffles.size===0');assert.equal(await evaluate('activeRaffle().revision'),1);
    // Several clicks can be registered while an earlier save is still in flight.
    await evaluate('window.__qa.hold=true;window.__qa.calls=[]');
    for(const n of ['20','21','22']){await prepare(n,'efectivo');await click('#saveNumberButton');}
    const pendingBatch=await evaluate('({pending:[...state.pending.keys()]})');
    assert.equal(pendingBatch.pending.length,3);assert.equal(await evaluate('state.summary.recaudado'),80000);
    await evaluate('window.__qa.hold=false;window.__qa.releases.splice(0).forEach(resolve=>resolve())');
    await wait('state.pending.size===0');assert(await evaluate('window.__qa.calls.some(c=>c.action==="batchNumbers"&&c.payload.changes.length===2)'));
    assert.equal(await evaluate('state.summary.recaudado'),140000);await evaluate('window.__qa.delay=120');
    // A rejected prize edit rolls the ticket back instead of claiming success.
    await click('[data-edit-prize]');await set('#editPremio','Premio rechazado');await evaluate('window.__qa.rejectPrize=true;document.getElementById("prizeForm").requestSubmit()');await wait('state.pendingRaffles.size===0');
    assert.equal(await evaluate('document.querySelector(".ticket-prize").textContent'),'$800.000');
    await wait('!state.refreshing');await screenshot('desktop-admin');
    // A raffle password never exposes platform approvals or free creation of other cards.
    assert(await evaluate('state.role==="owner"&&!isPlatform()&&document.getElementById("developerPanel").hidden&&document.getElementById("newRaffleButton").hidden'));
    await click('#adminButton');await wait('state.online&&!isAdmin()');assert.equal(await evaluate('document.getElementById("kpiRecaudado").textContent'),'');
    // Timeout and retry preserve the same request ID and never show a false success.
    await evaluate('window.__qa.failAction="reserveNumber";window.__qa.calls=[]');await click('[data-number="11"]');await set('#clienteNombre','Reintento QA');await set('#clienteTelefono','3001234567');await click('#saveNumberButton');
    await wait('!document.getElementById("retryChanges").hidden');assert(await evaluate('!document.getElementById("noticeDialog").open'));
    await click('#retryChanges');await wait('document.getElementById("noticeDialog").open');
    const ids=await evaluate('window.__qa.calls.filter(c=>c.action==="reserveNumber").map(c=>c.payload.requestId)');assert.equal(ids.length,2);assert.equal(ids[0],ids[1]);await click('[data-close="noticeDialog"]');
    assert(!(await evaluate('localStorage.getItem("rifapro.public.v2")')).includes('3001234567'));
    // The main CTA scrolls locally; payment opens WhatsApp but never activates a card.
    await evaluate('window.__qa.calls=[]');await click('#chooseNumbers');
    assert.equal(await evaluate('document.activeElement.id'),'numbersTitle');
    assert.equal(await evaluate('window.__qa.calls.length'),0);
    assert(await evaluate('document.querySelector(".page-intro").hidden&&document.querySelector(".creator-entry").hidden&&document.getElementById("adminButton").hidden&&document.getElementById("developerAccessButton").hidden&&!document.getElementById("acquireCardButton").hidden'));
    await click('#acquireCardButton');
    await set('#titulo','Mi rifa independiente');await set('#premio','Una bicicleta');
    await set('#precio','20000');await evaluate('document.getElementById("precio").dispatchEvent(new Event("input",{bubbles:true}))');
    assert.equal(await evaluate('document.getElementById("precio").value'),'$20.000');
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});await delay(180);await screenshot('creation-step-one');
    await click('#nextCreateStep');assert(await evaluate('creationStep===2&&!document.getElementById("createStepTwo").hidden'));
    await click('#previousCreateStep');assert.equal(await evaluate('document.getElementById("titulo").value'),'Mi rifa independiente');await click('#nextCreateStep');
    for(const width of [320,375,390,430]){
      await send('Emulation.setDeviceMetricsOverride',{width,height:844,deviceScaleFactor:1,mobile:true});
      for(const date of ['', '2099-11-01']){
        await set('#fecha',date);
        assert(await evaluate('(()=>{const d=document.getElementById("raffleDialog"),b=d.querySelector(".create-body"),r=b.getBoundingClientRect();return document.documentElement.scrollWidth<=innerWidth&&d.scrollWidth<=d.clientWidth&&b.scrollWidth<=b.clientWidth&&[...d.querySelectorAll("input")].filter(i=>i.getBoundingClientRect().width).every(i=>i.getBoundingClientRect().right<=r.right)})()'),'creation overflow at '+width+' date '+date);
      }
      assert(await evaluate('document.getElementById("createRaffleButton").getBoundingClientRect().bottom<=innerHeight'),'mobile action stays visible');
    }
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});
    await evaluate('document.querySelector(".create-body").scrollTop=0');
    const touchPoint=await evaluate('(()=>{const r=document.querySelector(".create-header").getBoundingClientRect();return {x:r.left+80,y:r.bottom-20}})()');
    const lockedTop=await evaluate('document.body.style.top');
    await send('Input.dispatchTouchEvent',{type:'touchStart',touchPoints:[touchPoint]});
    for(let i=1;i<=6;i++)await send('Input.dispatchTouchEvent',{type:'touchMove',touchPoints:[{x:touchPoint.x,y:touchPoint.y-i*15}]});
    await send('Input.dispatchTouchEvent',{type:'touchEnd',touchPoints:[]});
    assert(await evaluate('document.querySelector(".create-body").scrollTop>0'),'header gestures scroll modal content');
    assert.equal(await evaluate('document.body.style.top'),lockedTop,'background stays locked');
    await evaluate('document.querySelector(".create-body").scrollTop=0');
    for(const [id,value] of [['titulo','Mi rifa independiente'],['premio','Una bicicleta'],['fecha','2099-11-01'],['responsable','Organizador QA'],['creatorPhone','3007654321'],['creatorPin','7931'],['creatorPinConfirm','7931']])await set('#'+id,value);
    assert(await evaluate('Array.from(document.querySelectorAll("#raffleDialog input")).every(el=>parseFloat(getComputedStyle(el).fontSize)>=16)'));
    await screenshot('create-card-mobile');
    await evaluate('document.querySelector(".create-body").scrollTop=document.querySelector(".create-body").scrollHeight');await screenshot('create-card-pin-mobile');
    await click('#createRaffleButton');await wait('document.getElementById("paymentDialog").open');
    assert(await evaluate('hasSession()&&!isAdmin()&&state.ownerOrder.status==="pending"'));
    const ownerLink=await evaluate('document.getElementById("privateAccessLink").value');assert(ownerLink.includes('?rifa=')&&!ownerLink.includes('#acceso='));
    assert.equal(await evaluate('document.getElementById("creatorPin").value'),'');
    assert(!(await evaluate('localStorage.getItem("rifapro.public.v2")')).includes('#acceso='));
    const whatsapp=await evaluate('document.getElementById("paidWhatsApp").href');assert(whatsapp.startsWith('https://wa.me/573246394689?text='));assert(decodeURIComponent(whatsapp).includes('comprobante'));
    await evaluate('window.__copied="";navigator.clipboard.writeText=async text=>{window.__copied=text};document.getElementById("copyPaymentKey").click()');assert.equal(await evaluate('window.__copied'),'3208759798');
    assert.match(await evaluate('document.getElementById("paymentReference").textContent'),/^RP-\d{5}$/);
    assert(await evaluate('document.getElementById("paymentBrandsImage").complete&&document.getElementById("paymentBrandsImage").naturalWidth>=1000'));
    await click('#downloadPayment');await wait('document.getElementById("qrDialog").open');
    await wait('!document.getElementById("qrImageFrame").hidden&&document.getElementById("paymentQrImage").naturalWidth===1170');
    await delay(180);
    await screenshot('original-payment-qr');
    // Capture the native share call without saving or sending anything externally.
    await evaluate('navigator.canShare=()=>true;navigator.share=async data=>{window.__sharedFile=data.files[0].name}');
    await click('#saveQrImage');assert.equal(await evaluate('window.__sharedFile'),'qr-elkin-rangel.png');
    await evaluate('window.__savedQrAsset=qrAsset;window.__savedQrReady=qrAssetReady;qrAsset=null;qrAssetReady=Promise.resolve(null);showPaymentQr()');
    assert(await evaluate('!document.getElementById("qrUnavailable").hidden&&document.getElementById("qrImageFrame").hidden'));
    await evaluate('qrAsset=window.__savedQrAsset;qrAssetReady=window.__savedQrReady');
    await click('[data-close="qrDialog"]');assert(await evaluate('document.getElementById("paymentDialog").open&&document.body.classList.contains("modal-locked")'));
    await screenshot('card-payment-desktop');
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});await screenshot('card-payment-mobile');
    assert(await evaluate('document.getElementById("paymentDialog").scrollWidth<=document.getElementById("paymentDialog").clientWidth'));
    for(const dark of [true,false]){
      await evaluate(`setTheme(${dark})`);await delay(180);
      const unreadable=await evaluate(`(()=>{
        const rgb=s=>s.match(/[\\d.]+/g)?.slice(0,3).map(Number)||[0,0,0];
        const lum=c=>c.map(v=>{v/=255;return v<=.04045?v/12.92:((v+.055)/1.055)**2.4}).reduce((n,v,i)=>n+v*[.2126,.7152,.0722][i],0);
        return [...document.querySelectorAll('.button')].filter(b=>b.getClientRects().length&&!b.disabled).map(b=>{
          let p=b,bg='';while(p){bg=getComputedStyle(p).backgroundColor;if(!bg.endsWith(', 0)')&&bg!=='transparent')break;p=p.parentElement;}
          const a=lum(rgb(getComputedStyle(b).color)),z=lum(rgb(bg));return {text:b.textContent.trim().slice(0,35),ratio:(Math.max(a,z)+.05)/(Math.min(a,z)+.05)};
        }).filter(r=>r.ratio<4.5);
      })()`);
      assert.deepEqual(unreadable,[],'button contrast in '+(dark?'dark':'light')+' mode');
    }
    await evaluate('document.getElementById("paidWhatsApp").addEventListener("click",e=>e.preventDefault(),{once:true});document.getElementById("paidWhatsApp").click()');await wait('state.ownerOrder.status==="reviewing"');assert(await evaluate('!isAdmin()'));
    await click('[data-close="paymentDialog"]');await click('#adminButton');await wait('!hasSession()');
    await send('Page.navigate',{url:'http://127.0.0.1:4173/RIFA-PRO.COM/'});await wait('typeof state!=="undefined"&&state.online');
    assert(await evaluate('document.getElementById("developerAccessButton").getBoundingClientRect().width>0&&document.getElementById("developerAccessButton").textContent.trim()===""'));
    await click('#developerAccessButton');
    await wait('document.getElementById("developerPanel")&&!document.getElementById("developerPanel").hidden');
    assert(await evaluate('document.querySelector(".workspace").hidden&&document.getElementById("ordersPanel").hidden'));
    await set('#developerPassword','4826');await click('#developerLoginButton');await wait('document.getElementById("developerError").textContent.includes("incorrecta")');
    assert(await evaluate('!hasSession()'));await set('#developerPassword','7462');await click('#developerLoginButton');
    await wait('document.querySelector("[data-approve-card]")!==null');assert(await evaluate('isPlatform()&&!isAdmin()&&document.getElementById("adminSummary").hidden'));
    assert.equal(await evaluate('document.querySelectorAll(".order-row").length'),2);
    await send('Emulation.setDeviceMetricsOverride',{width:1440,height:1100,deviceScaleFactor:1,mobile:false});
    await screenshot('developer-dashboard-desktop');
    await set('#developerFilter','pending');await evaluate('document.getElementById("developerFilter").dispatchEvent(new Event("change"))');await wait('document.querySelectorAll(".order-row").length===1');
    await click('[data-approve-card]');await click('#confirmAccept');await wait('document.getElementById("ordersList").textContent.includes("No hay cartones")');
    await set('#developerFilter','active');await evaluate('document.getElementById("developerFilter").dispatchEvent(new Event("change"))');await wait('document.querySelectorAll(".order-row").length===2');
    await click('[data-reset-pin]');await set('#setupPin','7932');await set('#setupPinConfirm','7932');await click('#savePinButton');
    await wait('!document.getElementById("pinSetupDialog").open&&!document.getElementById("refreshOrders").disabled');
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});
    assert(await evaluate('document.documentElement.scrollWidth<=innerWidth'));await screenshot('developer-dashboard-mobile');
    await click('#adminButton');await wait('!hasSession()');assert.equal(await evaluate('document.getElementById("ordersList").textContent'),'');
    await send('Page.navigate',{url:ownerLink});
    await wait('typeof state!=="undefined"&&state.online&&document.querySelectorAll(".numero").length===100');
    assert(await evaluate('!hasSession()&&document.getElementById("adminSummary").hidden&&document.getElementById("shareButton").hidden'));
    await click('#adminButton');assert(await evaluate('document.getElementById("adminDialog").open'));
    await set('#pinInput','4826');await click('#loginButton');await wait('document.getElementById("adminError").textContent.includes("incorrecto")');
    await set('#pinInput','7932');await click('#loginButton');
    await wait('isAdmin()&&state.role==="owner"&&!document.getElementById("adminSummary").hidden');
    assert(await evaluate('document.getElementById("adminButton").classList.contains("organizer-access")&&!document.getElementById("shareButton").hidden'));
    assert.equal(await evaluate('state.rifas.length'),1);assert.equal(await evaluate('activeRaffle().titulo'),'Mi rifa independiente');assert(await evaluate('document.getElementById("ordersPanel").hidden'));
    await evaluate('window.__copied="";Object.defineProperty(navigator,"share",{value:undefined,configurable:true});navigator.clipboard.writeText=async text=>{window.__copied=text}');
    await click('#shareButton');await delay(150);assert(!(await evaluate('window.__copied')).includes('#acceso='));assert((await evaluate('window.__copied')).includes('?rifa='));
    assert((await evaluate('document.getElementById("toast").textContent')).includes('Enlace copiado'));
    await prepare('30','efectivo');await click('#saveNumberButton');await wait('state.pending.size===0');
    assert.equal(await evaluate('state.summary.recaudado'),20000);
    if(await evaluate('document.getElementById("noticeDialog").open'))await click('[data-close="noticeDialog"]');
    await click('#adminButton');await wait('!hasSession()&&state.online');
    assert(await evaluate('document.getElementById("shareButton").hidden&&document.getElementById("pinInput").value===""'));
    assert.equal(await evaluate('state.rifas.length'),1);
    // Returning organizers enter only their PIN from the root, restoring saved payments.
    await send('Page.navigate',{url:'http://127.0.0.1:4173/RIFA-PRO.COM/'});
    await wait('typeof state!=="undefined"&&state.online&&state.rifas.length===0');
    assert(await evaluate('getComputedStyle(document.getElementById("ownerAccessButton")).animationName==="return-glow"'));
    await click('#ownerAccessButton');
    assert(await evaluate('document.getElementById("adminDialog").open&&!document.querySelector("#adminDialog input[type=url]")&&!document.getElementById("ownerLoginDialog")'));
    await set('#pinInput','7932');await click('#loginButton');
    await wait('isAdmin()&&state.summary?.recaudado===20000');
    assert.equal(await evaluate('activeRaffle().titulo'),'Mi rifa independiente');
    assert(await evaluate('document.querySelector(\'[data-number="30"]\').classList.contains("pagado")'));
    assert(await evaluate('window.__qa.calls.find(c=>c.action==="login").payload.rifaId===undefined'));
    await click('#adminButton');await wait('!hasSession()&&state.online');
    for(const width of [320,375,390,768,1024]){
      await send('Emulation.setDeviceMetricsOverride',{width,height:844,deviceScaleFactor:1,mobile:width<768});await delay(120);
      assert(await evaluate('document.documentElement.scrollWidth<=innerWidth'),'horizontal overflow at '+width);
      assert(await evaluate('Array.from(document.querySelectorAll("input,select")).every(el=>parseFloat(getComputedStyle(el).fontSize)>=16)'),'input fonts below 16px');
      if(width===390){await evaluate('document.getElementById("toast").hidden=true');await screenshot('mobile-public');}
    }
    await send('Emulation.setDeviceMetricsOverride',{width:390,height:844,deviceScaleFactor:1,mobile:true});await click('#themeToggle');await screenshot('mobile-dark');await click('#themeToggle');
    await wait('navigator.serviceWorker.controller!==null');
    assert(await evaluate('caches.keys().then(async keys=>Boolean(await (await caches.open(keys.find(k=>k.startsWith("rifapro-shell")))).match(new URL("./images/breb-nequi.png",location.href))))'));
    const assets=await evaluate('Promise.all(["./images/icon-192.png","./images/icon-512.png","./images/apple-touch-icon.png"].map(src=>new Promise(resolve=>{const i=new Image();i.onload=()=>resolve(i.naturalWidth);i.onerror=()=>resolve(0);i.src=src})))');assert.deepEqual(assets,[192,512,180]);
    await send('Network.enable');await send('Network.emulateNetworkConditions',{offline:true,latency:0,downloadThroughput:0,uploadThroughput:0});await wait('!navigator.onLine');
    await click('[data-number="12"]');await wait('document.getElementById("noticeDialog").open');assert((await evaluate('document.getElementById("noticeTitle").textContent')).includes('Conéctate'));
    await send('Page.reload');await wait('document.getElementById("dashboard")!==null');assert.equal(await evaluate('document.title'),'Rifa Pro · Tu próximo número de la suerte');
    await send('Network.emulateNetworkConditions',{offline:false,latency:0,downloadThroughput:-1,uploadThroughput:-1});
    await send('Page.navigate',{url:'http://127.0.0.1:4173/RIFA-PRO.COM/?legacy=1&rifa=boleta'});await wait('typeof state!=="undefined"&&state.online&&document.querySelectorAll(".numero").length===100');
    assert.deepEqual(await evaluate('window.__qa.calls.slice(0,3).map(c=>c.action)'),['getSnapshot','getRifas','getRifaNumbers']);
    assert(await evaluate('!state.capabilities.independentRaffles'));
    assert.equal(await evaluate('state.rifas.length'),1);
    await send('Page.navigate',{url:'http://127.0.0.1:4173/RIFA-PRO.COM/'});await wait('typeof state!=="undefined"&&state.online');
    assert.equal(await evaluate('state.rifas.length'),0);assert(await evaluate('document.getElementById("numeros").hidden&&document.getElementById("developerPanel").hidden'));
    await screenshot('public-landing');
    assert.equal(errors.length,0,JSON.stringify(errors));
    log('PASS: immediate feedback '+instant.ms.toFixed(1)+'ms; stable DOM; single-request refresh; privacy; conflicts; transfer/cash; prize edits and rollback; batch writes; retry identity; participant-only shared links; two-step creation; formatted currency; header touch gestures with background locked; original QR preview/native share/failure fallback; prefetched cached PNGs; readable day/night buttons; purple Nequi in dark mode; short references; WhatsApp receipt; separate developer approval; owner PIN restore; clipboard confirmation; 320–1024 layouts; 16px inputs; PWA offline.');
    console.log('Screenshots saved in .qa/. Fixture backend only; live deployment still needs validation.');
  } finally { if(ws?.readyState===1)ws.close();browser.kill();preview.kill(); }
})().catch((e)=>{log(e.stack);process.exitCode=1;});
