// Isolated browser service. No live requests or real administrator credentials.
exports.bootstrap = function () {
  const legacy = new URL(location.href).searchParams.has('legacy');
  const catalog = [{ rifaId:'boleta', titulo:'Rifa de $500.000', premio:'$500.000', cantidad:100, precio:20000, loteria:'Boyacá', fecha:'2026-10-31', hora:'17:00', responsable:'Arelys B. Mendoza B.', revision:0 }];
  const records = [{numeroId:'00',estado:'apartado',clienteNombre:'Ana Pérez',clienteTelefono:'3001234567',revision:1,updatedAt:'2026-10-20T21:01:02Z'}, {numeroId:'01',estado:'pagado',clienteNombre:'Carlos',clienteTelefono:'3001234567',paymentMethod:'transferencia',revision:1}, {numeroId:'02',estado:'pagado',clienteNombre:'Luisa',clienteTelefono:'3001234567',paymentMethod:'efectivo',revision:1}];
  const capabilities = {snapshot:true,editPrize:true,batchNumbers:true,writeSnapshot:true,independentRaffles:true,expiry:true};
  const orders=[],ownerSessions=new Map();
  window.__qa = { calls:[], delay:120, failNext:false, rejectPrize:false, hold:false, releases:[] };
  const publicRecord = (n) => ({numeroId:n.numeroId,estado:n.estado==='libre'?'libre':'apartado',clienteNombre:n.clienteNombre,updatedAt:n.updatedAt});
  const summary = () => {
    const paid=records.filter(n=>n.estado==='pagado'),reserved=records.filter(n=>n.estado==='apartado').length;
    return {libres:100-paid.length-reserved,apartados:reserved,pagados:paid.length,recaudado:paid.length*20000,meta:2000000,transferencia:paid.filter(n=>n.paymentMethod==='transferencia').length*20000,efectivo:paid.filter(n=>n.paymentMethod==='efectivo').length*20000,sinClasificar:0};
  };
  const append=Element.prototype.appendChild;
  Element.prototype.appendChild=function(node){
    if(node.tagName==='SCRIPT' && node.src.startsWith('https://script.google.com')){
      const url=new URL(node.src), action=url.searchParams.get('action'), callback=url.searchParams.get('callback');
      window.__qa.calls.push({action});
      const id=url.searchParams.get('rifaId'),publicCatalog=catalog.filter(r=>id?r.rifaId===id:!r.independiente);
      const data=action==='getRifas'?{data:publicCatalog}:action==='getSnapshot'?{rifas:publicCatalog,rifaId:publicCatalog[0]?.rifaId||null,data:publicCatalog[0]?.independiente?[]:records.map(publicRecord)}:{data:records.map(publicRecord)};
      setTimeout(()=>window[callback]?.(legacy&&action==='getSnapshot'?{ok:false,error:'Acción pública no disponible.'}:{ok:true,protocol:2,...(legacy?{}:{capabilities}),...data}),30);return node;
    }
    return append.call(this,node);
  };
  function edit(action,p){
    if(p.numeroId==='08')return {ok:false,code:'CONFLICT',record:{numeroId:'08',clienteNombre:'Marta',updatedAt:'2026-10-20T21:01:02Z'}};
    const index=records.findIndex(n=>n.numeroId===p.numeroId);
    const next=action==='releaseNumber'?{numeroId:p.numeroId,estado:'libre',clienteNombre:'',revision:(records[index]?.revision||0)+1}:{...p,estado:action==='reserveNumber'?'apartado':p.estado,revision:(records[index]?.revision||0)+1};
    if(index>=0)records[index]=next;else records.push(next);
    return {ok:true,record:action==='reserveNumber'?publicRecord(next):next,summary:action==='reserveNumber'?undefined:summary()};
  }
  const realFetch=window.fetch;
  window.fetch=async function(url,options){
    if(!String(url).startsWith('https://script.google.com'))return realFetch.call(this,url,options);
    if(!navigator.onLine)throw new TypeError('offline');
    const {action,payload:p,token}=JSON.parse(options.body);window.__qa.calls.push({action,payload:p});
    await new Promise(resolve=>setTimeout(resolve,window.__qa.delay));
    if(window.__qa.hold)await new Promise(resolve=>window.__qa.releases.push(resolve));
    if(window.__qa.failAction===action){window.__qa.failAction=null;throw new TypeError('network failure');}
    let result={ok:true};
    if(action==='login')result=p.pin==='4826'?{ok:true,token:'qa-token',expiresAt:Date.now()+1800000}:{ok:false,code:'INVALID_PIN',error:'Contraseña incorrecta.'};
    if(action==='adminSnapshot'){const owned=ownerSessions.get(token),rifas=owned?catalog.filter(r=>r.rifaId===owned):catalog;result={ok:true,protocol:2,capabilities,rifas,rifaId:rifas.find(r=>r.rifaId===p.rifaId)?.rifaId||rifas[0]?.rifaId||null,data:owned?[]:records,summary:owned?{libres:100,apartados:0,pagados:0,recaudado:0,meta:500000,transferencia:0,efectivo:0}:summary()};}
    if(['reserveNumber','updateNumber','releaseNumber'].includes(action))result=edit(action,p);
    if(action==='batchNumbers')result={ok:true,results:p.changes.map(change=>({requestId:change.payload.requestId,...edit(change.action,change.payload)}))};
    if(action==='updatePrize'){
      if(window.__qa.rejectPrize){window.__qa.rejectPrize=false;result={ok:false,code:'STALE',error:'El premio cambió en otra sesión.'};}
      else{const r=catalog.find(r=>r.rifaId===p.rifaId);Object.assign(r,{premio:p.premio,titulo:p.titulo,revision:r.revision+1});result={ok:true,rifa:r};}
    }
    if(action==='createRaffle'){const r={...p,precio:Number(p.precio),cantidad:Number(p.cantidad),revision:0};catalog.unshift(r);result={ok:true,rifaId:p.rifaId,rifa:r};}
    if(action==='deleteRifa'){const i=catalog.findIndex(r=>r.rifaId===p.rifaId);if(i>=0)catalog.splice(i,1);}
    if(action==='orderCard'){let o=orders.find(o=>o.rifaId===p.rifaId);if(!o){o={...p,status:'pending',amount:1000};orders.push(o);}result={ok:true,order:{...o,ownerKey:undefined}};}
    if(action==='ownerLogin'){const o=orders.find(o=>o.rifaId===p.rifaId&&o.ownerKey===p.ownerKey);if(o){const key='owner-'+o.rifaId;ownerSessions.set(key,o.rifaId);result={ok:true,token:key,role:'owner',expiresAt:Date.now()+1800000,rifaId:o.rifaId,order:{...o,ownerKey:undefined}};}else result={ok:false,code:'INVALID_KEY',error:'El enlace privado no corresponde.'};}
    if(['ownerStatus','reportPayment'].includes(action)){const o=orders.find(o=>o.rifaId===ownerSessions.get(token));if(action==='reportPayment'&&o.status==='pending')o.status='reviewing';result={ok:true,order:{...o,ownerKey:undefined}};}
    if(action==='listOrders')result={ok:true,orders:orders.map(o=>({...o,ownerKey:undefined}))};
    if(action==='approveCard'){const o=orders.find(o=>o.rifaId===p.rifaId);o.status='active';catalog.unshift({...o,ownerKey:undefined,independiente:true,precio:Number(o.precio),cantidad:Number(o.cantidad),revision:0});result={ok:true,order:{...o,ownerKey:undefined}};}
    if(action==='logout')ownerSessions.delete(p.sessionToken);
    return new Response(JSON.stringify(result),{headers:{'Content-Type':'application/json'}});
  };
};
