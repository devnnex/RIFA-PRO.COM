// Isolated browser service. Dummy credentials and fake payments only.
exports.bootstrap = function () {
  const legacy = new URL(location.href).searchParams.has('legacy');
  const initial = {
    catalog: [{ rifaId:'boleta', titulo:'Rifa de $500.000', premio:'$500.000', cantidad:100, precio:20000, loteria:'Boyacá', fecha:'2026-10-31', hora:'17:00', responsable:'Organizadora QA', revision:0 }],
    records: { boleta: [{numeroId:'00',estado:'apartado',clienteNombre:'Ana Pérez',clienteTelefono:'3001234567',revision:1,updatedAt:'2026-10-20T21:01:02Z'}, {numeroId:'01',estado:'pagado',clienteNombre:'Carlos',clienteTelefono:'3001234567',paymentMethod:'transferencia',revision:1}, {numeroId:'02',estado:'pagado',clienteNombre:'Luisa',clienteTelefono:'3001234567',paymentMethod:'efectivo',revision:1}] }, orders: []
  };
  const data = JSON.parse(sessionStorage.getItem('qa.service') || 'null') || initial;
  const { catalog, orders } = data, sessions = new Map();
  const persist = () => sessionStorage.setItem('qa.service', JSON.stringify(data));
  const capabilities = {snapshot:true,editPrize:true,batchNumbers:true,writeSnapshot:true,independentRaffles:true,expiry:true,developerPanel:true,ownerPin:true};
  window.__qa = { calls:[], delay:120, rejectPrize:false, hold:false, releases:[] };
  const records = (id) => data.records[id] ||= [];
  const publicRecord = (n) => ({numeroId:n.numeroId,estado:n.estado==='libre'?'libre':'apartado',clienteNombre:n.clienteNombre,updatedAt:n.updatedAt});
  const summary = (id) => {
    const r=catalog.find(r=>r.rifaId===id), list=records(id),paid=list.filter(n=>n.estado==='pagado'),reserved=list.filter(n=>n.estado==='apartado').length;
    return {libres:r.cantidad-paid.length-reserved,apartados:reserved,pagados:paid.length,recaudado:paid.length*r.precio,meta:r.cantidad*r.precio,transferencia:paid.filter(n=>n.paymentMethod==='transferencia').length*r.precio,efectivo:paid.filter(n=>n.paymentMethod==='efectivo').length*r.precio,sinClasificar:0};
  };
  const append=Element.prototype.appendChild;
  Element.prototype.appendChild=function(node){
    if(node.tagName==='SCRIPT' && node.src.startsWith('https://script.google.com')){
      const url=new URL(node.src), action=url.searchParams.get('action'), callback=url.searchParams.get('callback'),id=url.searchParams.get('rifaId');
      window.__qa.calls.push({action});
      const list=catalog.filter(r=>legacy?true:id&&r.rifaId===id);
      const result=action==='getRifas'?{data:list}:action==='getSnapshot'?{rifas:list,rifaId:list[0]?.rifaId||null,data:list[0]?records(list[0].rifaId).map(publicRecord):[]}:{data:records(id).map(publicRecord)};
      setTimeout(()=>window[callback]?.(legacy&&action==='getSnapshot'?{ok:false,error:'Acción pública no disponible.'}:{ok:true,protocol:2,...(legacy?{}:{capabilities}),...result}),new URL(location.href).searchParams.has('slow')?900:30);return node;
    }
    return append.call(this,node);
  };
  const orderView = o => ({...o,ownerKey:undefined,adminPin:undefined,hasPin:Boolean(o.adminPin)});
  const denied = () => ({ok:false,code:'FORBIDDEN',error:'Este acceso no permite gestionar otro cartón.'});
  function edit(action,p){
    if(p.numeroId==='08')return {ok:false,code:'CONFLICT',record:{numeroId:'08',clienteNombre:'Marta',updatedAt:'2026-10-20T21:01:02Z'}};
    const list=records(p.rifaId),index=list.findIndex(n=>n.numeroId===p.numeroId);
    const next=action==='releaseNumber'?{numeroId:p.numeroId,estado:'libre',clienteNombre:'',revision:(list[index]?.revision||0)+1}:{...p,estado:action==='reserveNumber'?'apartado':p.estado,revision:(list[index]?.revision||0)+1};
    if(index>=0)list[index]=next;else list.push(next);
    return {ok:true,record:action==='reserveNumber'?publicRecord(next):next,summary:action==='reserveNumber'?undefined:summary(p.rifaId)};
  }
  const realFetch=window.fetch;
  window.fetch=async function(url,options){
    if(!String(url).startsWith('https://script.google.com'))return realFetch.call(this,url,options);
    if(!navigator.onLine)throw new TypeError('offline');
    const {action,payload:p,token}=JSON.parse(options.body);window.__qa.calls.push({action,payload:p});
    await new Promise(resolve=>setTimeout(resolve,window.__qa.delay));
    if(window.__qa.hold)await new Promise(resolve=>window.__qa.releases.push(resolve));
    if(window.__qa.failAction===action){window.__qa.failAction=null;throw new TypeError('network failure');}
    let result={ok:true}; const actor=sessions.get(token);
    const ownerOrder = () => orders.find(o=>o.rifaId===actor?.rifaId);
    const session = (actor,order) => {const token=crypto.randomUUID();sessions.set(token,actor);return {ok:true,token,expiresAt:Date.now()+1800000,...actor,order};};
    if(action==='login'){
      const o=orders.find(o=>o.rifaId===p.rifaId);
      result=o?(o.adminPin===p.pin&&o.status==='active'?session({role:'owner',rifaId:o.rifaId},orderView(o)):{ok:false,code:'INVALID_PIN',error:'PIN incorrecto.'}):p.pin==='4826'&&p.rifaId==='boleta'?session({role:'owner',rifaId:'boleta',legacy:true},{rifaId:'boleta',status:'active',legacy:true,hasPin:true,titulo:catalog.find(r=>r.rifaId==='boleta').titulo}):{ok:false,code:'INVALID_PIN',error:'PIN incorrecto.'};
    }
    if(action==='developerLogin')result=p.password==='Developer-fixture-4826'?session({role:'developer'}):{ok:false,code:'INVALID_PIN',error:'Contraseña incorrecta.'};
    if(['developerDashboard','listOrders','approveCard','createRaffle'].includes(action)&&actor?.role!=='developer')return new Response(JSON.stringify(denied()));
    if(['adminSnapshot','updateNumber','releaseNumber','updatePrize','deleteRifa'].includes(action)&&actor?.role!=='developer'&&(!actor||p.rifaId&&p.rifaId!==actor.rifaId))return new Response(JSON.stringify(denied()));
    if(action==='adminSnapshot'){const rifas=actor.role==='developer'?catalog:catalog.filter(r=>r.rifaId===actor.rifaId),id=rifas.find(r=>r.rifaId===p.rifaId)?.rifaId||rifas[0]?.rifaId;result={ok:true,protocol:2,capabilities,rifas,rifaId:id,data:records(id),summary:summary(id)};}
    if(['reserveNumber','updateNumber','releaseNumber'].includes(action))result=edit(action,p);
    if(action==='batchNumbers')result={ok:true,results:p.changes.map(change=>({requestId:change.payload.requestId,...(actor?.rifaId===change.payload.rifaId?edit(change.action,change.payload):denied())}))};
    if(action==='updatePrize'){
      if(window.__qa.rejectPrize){window.__qa.rejectPrize=false;result={ok:false,code:'STALE',error:'El premio cambió en otra sesión.'};}
      else{const r=catalog.find(r=>r.rifaId===p.rifaId);Object.assign(r,{premio:p.premio,titulo:p.titulo,revision:r.revision+1});result={ok:true,rifa:r};}
    }
    if(action==='deleteRifa'){const i=catalog.findIndex(r=>r.rifaId===p.rifaId);if(i>=0)catalog.splice(i,1);const o=orders.find(o=>o.rifaId===p.rifaId);if(o)o.status='deleted';}
    if(action==='orderCard'){let o=orders.find(o=>o.rifaId===p.rifaId);if(!o){o={...p,status:'pending',amount:1000};orders.push(o);}result={ok:true,order:orderView(o)};}
    if(action==='ownerLogin'){const o=orders.find(o=>o.rifaId===p.rifaId&&o.ownerKey===p.ownerKey);result=o?session({role:'owner',rifaId:o.rifaId},orderView(o)):{ok:false,code:'INVALID_KEY',error:'El enlace privado no corresponde.'};}
    if(action==='setOwnerPin'){const o=ownerOrder();if(actor.rifaId!==p.rifaId)result=denied();else{o.adminPin=p.pin;result={ok:true,order:orderView(o)};}}
    if(['ownerStatus','reportPayment'].includes(action)){const o=ownerOrder();if(action==='reportPayment'&&o.status==='pending')o.status='reviewing';result={ok:true,order:actor.legacy?{rifaId:actor.rifaId,status:'active',legacy:true}:orderView(o)};}
    if(action==='developerDashboard')result={ok:true,capabilities,cards:[...orders.map(orderView),...catalog.filter(r=>!orders.some(o=>o.rifaId===r.rifaId)).map(r=>({...r,status:'active',legacy:true}))]};
    if(action==='listOrders')result={ok:true,orders:orders.map(orderView)};
    if(action==='approveCard'){const o=orders.find(o=>o.rifaId===p.rifaId);o.status='active';if(!catalog.some(r=>r.rifaId===o.rifaId))catalog.unshift({...orderView(o),independiente:true,precio:Number(o.precio),cantidad:Number(o.cantidad),revision:0});result={ok:true,order:orderView(o)};}
    if(action==='logout')sessions.delete(p.sessionToken);
    persist();return new Response(JSON.stringify(result),{headers:{'Content-Type':'application/json'}});
  };
};
