const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const crypto = require('node:crypto');
const code = fs.readFileSync('backend/Code.gs', 'utf8');

function server(legacyRows) {
  const sheets = new Map(), properties = new Map([['SPREADSHEET_ID','test'],['ADMIN_PIN','4826'],['DEVELOPER_PIN','7462']]), cache = new Map();
  let locked = false;
  const metrics = { entryReads: 0, entryWrites: 0, ledgerWrites: 0, opens: 0 };
  class Sheet {
    constructor() { this.data = []; }
    appendRow(row) { this.data.push([...row]); }
    setFrozenRows() {}
    getLastRow() { return this.data.length; }
    deleteRow(row) { assert(locked); this.data.splice(row - 1, 1); }
    getRange(start, column, count, width) {
      return { getValues: () => { if(this.name==='rifas' && start===2)metrics.entryReads++;return Array.from({length:count},(_,i) => Array.from({length:width},(_,j) => this.data[start-1+i]?.[column-1+j] ?? '')); },
        setNumberFormat() {}, setValues: (values) => { assert(locked, 'writes require the server lock'); assert.equal(values.length,count); for(const row of values)assert.equal(row.length,width); if(this.name==='rifas')metrics.entryWrites++;else metrics.ledgerWrites++;values.forEach((r,i) => { this.data[start-1+i] ||= []; r.forEach((v,j) => {this.data[start-1+i][column-1+j]=v;}); }); } };
    }
  }
  const db = { getId: () => 'test', getSheetByName: (name) => sheets.get(name), insertSheet: (name) => { const s = new Sheet();s.name=name;sheets.set(name,s); return s; } };
  if (legacyRows) { const sheet = new Sheet();sheet.name='rifas';sheet.data = legacyRows.map(row=>[...row]); sheets.set('rifas',sheet); }
  const ctx = vm.createContext({ console, Date, JSON, Math, String, Number, Error,
    SpreadsheetApp: { openById: () => {metrics.opens++;return db;}, getActiveSpreadsheet: () => db, flush() {} },
    LockService: { getScriptLock: () => ({ tryLock: () => { assert(!locked); locked = true; return true; }, releaseLock: () => { locked = false; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (key) => properties.get(key), setProperty: (k,v) => properties.set(k,v), deleteProperty: (key) => properties.delete(key) }) },
    CacheService: { getScriptCache: () => ({ get: (key) => cache.get(key), put: (k,v) => cache.set(k,v), remove: (key) => cache.delete(key) }) },
    Utilities: { getUuid: () => crypto.randomUUID(), DigestAlgorithm: {SHA_256:'sha256'}, Charset: {UTF_8:'utf8'}, computeDigest: (algorithm,text) => [...crypto.createHash(algorithm).update(text).digest()], formatDate: (value,zone,format) => { const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:zone,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).formatToParts(value).map(p=>[p.type,p.value])); return format==='HH:mm'?`${parts.hour}:${parts.minute}`:`${parts.year}-${parts.month}-${parts.day}`; } },
    ContentService: { MimeType: { JSON:'json', JAVASCRIPT:'js' }, createTextOutput: (value) => ({ value, setMimeType() { return this; } }) }
  });
  vm.runInContext(code,ctx);
  ctx.prepararSistema();
  if (!legacyRows) ctx.createRaffle_({ rifaId:'test',titulo:'Rifa',cantidad:100,precio:20000,premio:'$500.000',loteria:'Boyacá',fecha:'2026-10-31',hora:'17:00',responsable:'Organizadora' });
  ctx.prepararSistema();
  const post = (action,payload = {}, token) => JSON.parse(ctx.doPost({ postData:{contents:JSON.stringify({action,payload,token})} }).value);
  const get = (action,params = {}) => JSON.parse(ctx.doGet({ parameter:{action,...params} }).value);
  const login = () => post('developerLogin',{password:'7462'}).token;
  return {ctx,post,get,login,cache,properties,sheets,metrics};
}
function reservation(extra = {}) { return { rifaId:'test',numeroId:'07',clienteNombre:'Ana Pérez',clienteTelefono:'3001234567',requestId:crypto.randomUUID(),...extra }; }

test('first confirmed reservation wins, even inside the same second', () => {
  const s = server(); const a = s.post('reserveNumber',reservation()); const b = s.post('reserveNumber',reservation({clienteNombre:'Luis'}));
  assert.equal(a.ok,true); assert.equal(b.code,'CONFLICT'); assert.equal(b.record.clienteNombre,'Ana Pérez');
  assert.equal(s.get('getRifaNumbers',{rifaId:'test'}).data.length,1);
});
test('simultaneous independent client requests yield one winner and preserve its name', async () => {
  const s = server();
  const results = await Promise.all(Array.from({length:20}, (_,i) => Promise.resolve().then(() => s.post('reserveNumber',reservation({clienteNombre:`Cliente ${i}`})))));
  assert.equal(results.filter((r) => r.ok).length,1);
  assert.equal(results.filter((r) => r.code === 'CONFLICT').length,19);
  assert(results.slice(1).every((r) => r.record.clienteNombre === 'Cliente 0'));
});
test('retry after uncertain network result is idempotent', () => {
  const s = server(), p = reservation(); assert.deepEqual(s.post('reserveNumber',p),s.post('reserveNumber',p));
  assert.equal(s.get('getRifaNumbers',{rifaId:'test'}).data.length,1);
});
test('public callers cannot spoof paid status or a payment method', () => {
  const s = server(); s.post('reserveNumber',reservation({estado:'pagado',paymentMethod:'efectivo'}));
  const r = s.post('adminSnapshot',{rifaId:'test'},s.login()); assert.equal(r.data[0].estado,'apartado'); assert.equal(r.summary.recaudado,0);
});
test('private read, edits, deletion and creation require a server session', () => {
  const s = server();
  for (const action of ['adminSnapshot','updateNumber','releaseNumber','deleteRifa','createRaffle','updatePrize','batchNumbers']) assert.equal(s.post(action,reservation()).code,'UNAUTHORIZED');
  assert.equal(s.get('verifyPin').ok,false); assert.equal(s.get('upsertNumber').ok,false);
});
test('public snapshot shows paid cash while hiding phones, totals and write metadata', () => {
  const s = server(), token = s.login(); s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'efectivo',revision:0}),token);
  const pub = s.get('getRifaNumbers',{rifaId:'test'}); assert.equal(pub.data[0].estado,'pagado');assert.equal(pub.data[0].paymentMethod,'efectivo');
  const json = JSON.stringify(pub); for (const secret of ['3001234567','recaudado','revision','requestId']) assert(!json.includes(secret),secret);
});
test('Configured PIN works; incorrect PIN fails; logout revokes access', () => {
  const s = server(); assert.equal(s.post('login',{pin:'9999'}).ok,false); const token = s.login(); assert(token);
  assert.equal(s.post('adminSnapshot',{rifaId:'test'},token).ok,true); s.post('logout',{sessionToken:token});
  assert.equal(s.post('adminSnapshot',{rifaId:'test'},token).code,'UNAUTHORIZED');
});
test('repeated wrong PINs are throttled across client identities', () => {
  const s = server(); for (let i=0;i<8;i++) s.post('login',{pin:'9999',rifaId:'test',clientId:crypto.randomUUID()});
  assert.equal(s.post('login',{pin:'4826',rifaId:'test'}).code,'RATE_LIMIT');
});
test('paid requires a valid method; totals split by transfer and cash', () => {
  const s = server(), token = s.login();
  assert.equal(s.post('updateNumber',reservation({estado:'pagado',revision:0}),token).ok,false);
  assert.equal(s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'transferencia',revision:0}),token).ok,true);
  assert.equal(s.post('updateNumber',reservation({numeroId:'08',estado:'pagado',paymentMethod:'efectivo',revision:0}),token).ok,true);
  const snap = s.post('adminSnapshot',{rifaId:'test'},token);
  assert.equal(snap.summary.recaudado,40000); assert.equal(snap.summary.transferencia,20000); assert.equal(snap.summary.efectivo,20000); assert.equal(snap.summary.sinClasificar,0);
});
test('changing paid to reserved clears payment and decreases revenue', () => {
  const s = server(), token = s.login(); s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'efectivo',revision:0}),token);
  s.post('updateNumber',reservation({estado:'apartado',paymentMethod:'efectivo',revision:1}),token);
  const snap = s.post('adminSnapshot',{rifaId:'test'},token); assert.equal(snap.summary.recaudado,0); assert.equal(snap.data[0].paymentMethod,'');
});
test('stale admin edits and releases cannot overwrite a newer reservation', () => {
  const s = server(), token = s.login(); s.post('reserveNumber',reservation());
  assert.equal(s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'efectivo',revision:0}),token).code,'STALE');
  assert.equal(s.post('releaseNumber',reservation({revision:0}),token).code,'STALE');
});
test('release removes customer, phone and payment, and permits a new reservation', () => {
  const s = server(), token = s.login(); s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'efectivo',revision:0}),token);
  assert.equal(s.post('releaseNumber',reservation({revision:1}),token).ok,true);
  const snap = s.post('adminSnapshot',{rifaId:'test'},token); assert.equal(snap.data[0].clienteNombre,''); assert.equal(snap.data[0].clienteTelefono,''); assert.equal(snap.summary.libres,100);
  assert.equal(s.post('reserveNumber',reservation({clienteNombre:'Luis'})).ok,true);
});
test('old reservation retry cannot overwrite a reassigned number', () => {
  const s = server(), token = s.login(), original = reservation(); s.post('reserveNumber',original);
  s.post('releaseNumber',reservation({revision:1}),token); s.post('reserveNumber',reservation({clienteNombre:'Luis'}));
  assert.equal(s.post('reserveNumber',original).code,'CONFLICT'); assert.equal(s.get('getRifaNumbers',{rifaId:'test'}).data[0].clienteNombre,'Luis');
});
test('bounds, customer validation and malformed request IDs are rejected', () => {
  const s = server(); for (const extra of [{numeroId:'100'},{numeroId:'-1'},{numeroId:'NaN'},{clienteNombre:' '},{clienteTelefono:'123'},{requestId:'x'}]) assert.equal(s.post('reserveNumber',reservation(extra)).ok,false);
});
test('duplicate raffle creation is idempotent and deletion isolates other raffles', () => {
  const s = server(), token = s.login(); assert.equal(s.post('createRaffle',{rifaId:'test'},token).ok,true);
  assert.equal(s.get('getRifas',{rifaId:'test'}).data.length,1); s.post('reserveNumber',reservation());
  assert.equal(s.post('deleteRifa',{rifaId:'test'},token).ok,true); assert.equal(s.get('getRifas',{rifaId:'test'}).data.length,0);
  assert.equal(s.post('deleteRifa',{rifaId:'test'},token).ok,true);
  assert.equal(s.post('reserveNumber',reservation()).code,'NOT_FOUND');
});
test('JSONP rejects executable callback names and exposes read actions only', () => {
  const s = server(); const out = s.ctx.doGet({parameter:{action:'getRifas',callback:'alert(1)//'}}).value;
  assert.equal(JSON.parse(out).ok,false); assert(!out.includes('alert(1)'));
  const valid = s.ctx.doGet({parameter:{action:'getRifas',callback:'rifa_valid123'}}).value; assert(valid.startsWith('rifa_valid123('));
});
test('setup is repeatable and does not delete existing data', () => {
  const s = server(); s.post('reserveNumber',reservation()); s.ctx.prepararSistema();
  assert.equal(s.get('getRifas',{rifaId:'test'}).data.length,1); assert.equal(s.get('getRifaNumbers',{rifaId:'test'}).data.length,1);
});

const legacyHeaders=['rifaId','numeroId','estado','clienteNombre','clienteTelefono','updatedAt','titulo','precio','premio','loteria','fecha','hora','responsable'];
const legacyMeta=['old','__META__','meta','100','','2026-10-01T10:00:00Z','Rifa anterior',20000,'$500.000','Boyacá','2026-10-31','17:00','Organizadora'];
const legacyPaid=['old','07','pagado','Ana anterior','3007654321','2026-10-02T10:00:00Z','Rifa anterior',20000,'$500.000','Boyacá','2026-10-31','17:00','Organizadora'];
test('original 13-column spreadsheet and __META__ rows remain readable without migration', () => {
  const s=server([legacyHeaders,legacyMeta,legacyPaid]);
  const r=s.get('getRifas',{rifaId:'old'});assert.equal(r.data[0].titulo,'Rifa anterior');assert.equal(r.data[0].cantidad,100);
  const snap=s.post('adminSnapshot',{rifaId:'old'},s.login());assert.equal(snap.data[0].clienteNombre,'Ana anterior');assert.equal(snap.data[0].clienteTelefono,'3007654321');
  assert.equal(snap.summary.recaudado,20000);assert.equal(snap.summary.sinClasificar,20000);
  assert.deepEqual(s.sheets.get('rifas').data[1],legacyMeta);assert.deepEqual(s.sheets.get('rifas').data[2],legacyPaid);
  assert.deepEqual(s.sheets.get('rifas').data[0].slice(0,13),legacyHeaders);
});
test('an existing paid number stays unavailable and its phone stays private', () => {
  const s=server([legacyHeaders,legacyMeta,legacyPaid]);
  const result=s.post('reserveNumber',reservation({rifaId:'old'}));assert.equal(result.code,'CONFLICT');assert.equal(result.record.clienteNombre,'Ana anterior');
  const pub=s.get('getRifaNumbers',{rifaId:'old'});assert.equal(pub.data[0].estado,'pagado');assert(!JSON.stringify(pub).includes('3007654321'));
});
test('classifying an old paid record preserves the original sheet layout and customer', () => {
  const s=server([legacyHeaders,legacyMeta,legacyPaid]); const token=s.login();
  const result=s.post('updateNumber',reservation({rifaId:'old',clienteNombre:'Ana anterior',clienteTelefono:'3007654321',estado:'pagado',paymentMethod:'transferencia',revision:0}),token);
  assert.equal(result.ok,true);const row=s.sheets.get('rifas').data[2];assert.equal(row[5].length>0,true);assert.equal(row[13],'transferencia');assert.equal(row[14],1);assert.equal(row[6],'Rifa anterior');
  const summary=s.post('adminSnapshot',{rifaId:'old'},token).summary;assert.equal(summary.transferencia,20000);assert.equal(summary.sinClasificar,0);
});
test('bound spreadsheet is supported and setup saves its ID for the web application', () => {
  const s=server();s.properties.delete('SPREADSHEET_ID');s.ctx.prepararSistema();assert.equal(s.properties.get('SPREADSHEET_ID'),'test');
});
test('date and time cells from Sheets are normalized for the existing ticket', () => {
  const meta=[...legacyMeta];meta[10]=new Date('2026-10-31T05:00:00Z');meta[11]=new Date('2026-10-31T22:00:00Z');
  const s=server([legacyHeaders,meta]);const r=s.get('getRifas',{rifaId:'old'}).data[0];assert.equal(r.fecha,'2026-10-31');assert.equal(r.hora,'17:00');
});
test('unrelated raffles and their rows survive deletion', () => {
  const other=[...legacyMeta];other[0]='other';const s=server([legacyHeaders,legacyMeta,legacyPaid,other]);
  assert.equal(s.post('deleteRifa',{rifaId:'old'},s.login()).ok,true);assert.equal(s.get('getRifas',{rifaId:'other'}).data[0].rifaId,'other');
});
test('health endpoint is preserved', () => { assert.equal(server().get('health').sheet,'rifas'); });
test('editing a duplicated legacy number cleans only its duplicates', () => {
  const duplicate=[...legacyPaid];duplicate[1]='7';duplicate[3]='Nombre actualizado';duplicate[5]='2026-10-03T10:00:00Z';
  const s=server([legacyHeaders,legacyMeta,legacyPaid,duplicate]);assert.equal(s.post('adminSnapshot',{rifaId:'old'},s.login()).summary.pagados,1);
  const result=s.post('updateNumber',reservation({rifaId:'old',estado:'pagado',paymentMethod:'efectivo',revision:0}),s.login());assert.equal(result.ok,true);
  assert.equal(s.sheets.get('rifas').data.length,3);assert.equal(s.get('getRifaNumbers',{rifaId:'old'}).data.length,1);
});
test('unexpected custom headers cause setup to stop without overwriting customer data', () => {
  const s=server();s.sheets.get('rifas').data[0][13]='custom';const snapshot=JSON.stringify(s.sheets.get('rifas').data);
  assert.throws(()=>s.ctx.prepararSistema());assert.equal(JSON.stringify(s.sheets.get('rifas').data),snapshot);
});
test('one public snapshot includes paid status and method without totals or contact fields', () => {
  const s=server();s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'efectivo',revision:0}),s.login());
  const snap=s.get('getSnapshot',{rifaId:'test'});assert.equal(snap.rifas[0].rifaId,'test');assert.equal(snap.data[0].estado,'pagado');assert.equal(snap.data[0].paymentMethod,'efectivo');assert.equal(snap.capabilities.snapshot,true);
  for(const key of ['clienteTelefono','summary'])assert(!JSON.stringify(snap).includes(key));
});
test('an administrator snapshot reads the entries once and opens the database once', () => {
  const s=server(),token=s.login();Object.keys(s.metrics).forEach(key=>s.metrics[key]=0);
  const snap=s.post('adminSnapshot',{rifaId:'test'},token);assert.equal(snap.rifas.length,1);assert.equal(s.metrics.entryReads,1);assert.equal(s.metrics.opens,1);
});
test('number write returns confirmed record and summary without a follow-up request', () => {
  const s=server();s.ctx.SpreadsheetApp.flush=()=>s.cache.set('public:v3:test',JSON.stringify([{numeroId:'07',estado:'libre'}]));
  const result=s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'transferencia',revision:0}),s.login());
  assert.equal(result.record.paymentMethod,'transferencia');assert.equal(result.record.revision,1);assert.equal(result.summary.recaudado,20000);
  assert.equal(s.get('getRifaNumbers',{rifaId:'test'}).data[0].estado,'pagado');
});
test('editing the prize keeps participants, payments, price and quantity unchanged', () => {
  const s=server([legacyHeaders,legacyMeta,legacyPaid]),token=s.login();
  const result=s.post('updatePrize',{rifaId:'old',premio:'Una bicicleta',titulo:'Rifa de una bicicleta',revision:0,requestId:crypto.randomUUID()},token);
  assert.equal(result.ok,true);assert.equal(result.rifa.premio,'Una bicicleta');
  const snap=s.post('adminSnapshot',{rifaId:'old'},token);assert.equal(snap.data[0].clienteNombre,'Ana anterior');assert.equal(snap.summary.recaudado,20000);assert.equal(snap.rifas[0].precio,20000);assert.equal(snap.rifas[0].cantidad,100);
  assert.equal(s.get('getSnapshot',{rifaId:'old'}).rifas[0].premio,'Una bicicleta');
});
test('prize edits are idempotent, reject stale changes and invalidate cached metadata', () => {
  const s=server(),token=s.login(),p={rifaId:'test',premio:'$800.000',titulo:'Rifa nueva',revision:0,requestId:crypto.randomUUID()};s.get('getRifas',{rifaId:'test'});
  assert.equal(s.post('updatePrize',p,token).ok,true);assert.equal(s.post('updatePrize',p,token).ok,true);
  assert.equal(s.post('updatePrize',{...p,requestId:crypto.randomUUID(),premio:'Otro'},token).code,'STALE');
  assert.equal(s.get('getRifas',{rifaId:'test'}).data[0].premio,'$800.000');assert.equal(s.get('getRifas',{rifaId:'test'}).data[0].revision,1);
});
test('batch writes group contiguous number rows and the ledger while reading only once', () => {
  const s=server(),token=s.login();Object.keys(s.metrics).forEach(key=>s.metrics[key]=0);
  const changes=Array.from({length:10},(_,i)=>({action:'updateNumber',payload:reservation({numeroId:String(i).padStart(2,'0'),estado:'pagado',paymentMethod:i%2?'efectivo':'transferencia',revision:0})}));
  const result=s.post('batchNumbers',{changes},token);assert.equal(result.results.filter(r=>r.ok).length,10);assert.equal(s.metrics.entryReads,1);assert.equal(s.metrics.entryWrites,1);assert.equal(s.metrics.ledgerWrites,1);
  const snap=s.post('adminSnapshot',{rifaId:'test'},token);assert.equal(snap.summary.recaudado,200000);assert.equal(snap.summary.transferencia,100000);assert.equal(snap.summary.efectivo,100000);
  assert.equal(s.post('batchNumbers',{changes},token).results.filter(r=>r.ok).length,10);assert.equal(s.get('getRifaNumbers',{rifaId:'test'}).data.length,10);
});
test('batch validation isolates invalid changes and prevents stale overwrites', () => {
  const s=server(),token=s.login();s.post('reserveNumber',reservation());
  const result=s.post('batchNumbers',{changes:[{action:'updateNumber',payload:reservation({estado:'pagado',paymentMethod:'efectivo',revision:0})},{action:'updateNumber',payload:reservation({numeroId:'08',estado:'pagado',paymentMethod:'efectivo',revision:0})}]},token);
  assert.equal(result.results[0].code,'STALE');assert.equal(result.results[1].ok,true);assert.equal(s.post('adminSnapshot',{rifaId:'test'},token).summary.recaudado,20000);
  assert.equal(s.post('batchNumbers',{changes:[]},token).ok,false);
});
test('batch updates retain the cleanup behavior for duplicated legacy rows', () => {
  const duplicate=[...legacyPaid];duplicate[1]='7';const s=server([legacyHeaders,legacyMeta,legacyPaid,duplicate]);
  const result=s.post('batchNumbers',{changes:[{action:'updateNumber',payload:reservation({rifaId:'old',estado:'pagado',paymentMethod:'efectivo',revision:0})}]},s.login());
  assert.equal(result.results[0].ok,true);assert.equal(s.sheets.get('rifas').data.length,3);
});

let cardSequence = 1000;
function card(extra={}) { return {rifaId:crypto.randomUUID(),ownerKey:crypto.randomUUID()+crypto.randomUUID(),adminPin:String(++cardSequence),contact:'3007654321',titulo:'Rifa independiente',premio:'Una bicicleta',cantidad:100,precio:5000,loteria:'Boyacá',fecha:'2099-11-01',hora:'17:00',responsable:'Organizador',...extra}; }
function ordered(s,p=card()) { const order=s.post('orderCard',p);assert.equal(order.ok,true,JSON.stringify(order));const session=s.post('ownerLogin',{rifaId:p.rifaId,ownerKey:p.ownerKey});assert.equal(session.ok,true);return {p,token:session.token,session}; }

test('winner publication requires the raffle owner and cannot cross card boundaries',()=>{
  const s=server(),a=ordered(s),b=ordered(s),dev=s.login();
  const p=reservation({rifaId:a.p.rifaId,revision:0});
  assert.equal(s.post('publishWinner',p).code,'UNAUTHORIZED');
  assert.equal(s.post('publishWinner',p,a.token).code,'PAYMENT_PENDING');
  s.post('approveCard',{rifaId:a.p.rifaId},dev); s.post('approveCard',{rifaId:b.p.rifaId},dev);
  s.post('reserveNumber',p);
  assert.equal(s.post('publishWinner',p,b.token).code,'FORBIDDEN');
  assert.equal(s.get('publishWinner',p).ok,false);
  assert.equal(s.post('publishWinner',p,a.token).ok,true);
  assert.equal(s.get('getSnapshot',{rifaId:b.p.rifaId}).rifas[0].winner,null);
});
test('winner is derived from the participant, published on its link and retries are idempotent',()=>{
  const s=server(),token=s.post('login',{pin:'4826',rifaId:'test'}).token;
  s.post('reserveNumber',reservation());
  s.get('getSnapshot',{rifaId:'test'}); // Populate catalog cache before publication.
  const p=reservation({revision:0,nombre:'Spoofed name'}),first=s.post('publishWinner',p,token);
  assert.equal(first.ok,true);assert.equal(first.winner.numeroId,'07');assert.equal(first.winner.nombre,'Ana Pérez');
  assert.deepEqual(s.post('publishWinner',p,token),first);
  assert.equal(s.sheets.get('RifaPro_Ganadores').data.length,2);
  const pub=s.get('getSnapshot',{rifaId:'test'});assert.deepEqual(pub.rifas[0].winner,first.winner);
  assert(!JSON.stringify(pub).includes('3001234567'));assert(!JSON.stringify(pub).includes(p.requestId));
  assert.deepEqual(s.get('getSnapshot').rifas,[]);
});
test('winner rejects free numbers, invalid numbers and stale changes without touching payments',()=>{
  const s=server(),token=s.login();
  assert.equal(s.post('publishWinner',reservation({revision:0}),token).ok,false);
  assert.equal(s.post('publishWinner',reservation({numeroId:'999',revision:0}),token).ok,false);
  s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'efectivo',revision:0}),token);
  const before=s.post('adminSnapshot',{rifaId:'test'},token);
  assert.equal(s.post('publishWinner',reservation({revision:0}),token).ok,true);
  assert.equal(s.post('publishWinner',reservation({revision:0}),token).code,'STALE');
  const after=s.post('adminSnapshot',{rifaId:'test'},token);
  assert.deepEqual(after.data,before.data);assert.deepEqual(after.summary,before.summary);
});
test('winner may be published after expiry; deletion removes only that raffle result',()=>{
  const s=server(),token=s.post('login',{pin:'4826',rifaId:'test'}).token;
  s.post('reserveNumber',reservation());
  const entries=s.sheets.get('rifas').data;entries.slice(1).forEach(row=>row[10]='2000-01-01');
  assert.equal(s.post('publishWinner',reservation({revision:0}),token).ok,true);
  assert.equal(s.get('getSnapshot',{rifaId:'test'}).rifas[0].winner.nombre,'Ana Pérez');
  assert.equal(s.post('deleteRifa',{rifaId:'test'},token).ok,true);
  assert.equal(s.sheets.get('RifaPro_Ganadores').data.length,1);
});
test('existing raffles still load before winner table migration, which preserves original columns',()=>{
  const s=server();s.sheets.delete('RifaPro_Ganadores');
  assert.equal(s.get('getSnapshot',{rifaId:'test'}).ok,true);
  s.ctx.prepararSistema();assert(s.sheets.has('RifaPro_Ganadores'));
  assert.equal(s.sheets.get('rifas').data[0].length,17);
});
test('previous orders keep their quoted fee while new cards cost 3500',()=>{
  const s=server(),a=ordered(s),dev=s.login();
  s.sheets.get('RifaPro_Cartones').data[1][7]=1000;
  assert.equal(s.post('ownerStatus',{},a.token).order.amount,1000);
  assert.equal(s.post('approveCard',{rifaId:a.p.rifaId},dev).ok,true);
  const b=ordered(s);assert.equal(b.session.order.amount,3500);
});
test('developer deletes pending and reviewing orders before a raffle exists; retries cannot reactivate them',()=>{
  const s=server(),dev=s.login(),a=ordered(s),b=ordered(s);
  s.post('reportPayment',{},b.token);
  assert.equal(s.post('deleteRifa',{rifaId:a.p.rifaId},b.token).code,'FORBIDDEN');
  assert.equal(s.post('deleteRifa',{rifaId:a.p.rifaId},a.token).code,'PAYMENT_PENDING');
  for(const item of [a,b]) {
    assert.equal(s.post('deleteRifa',{rifaId:item.p.rifaId},dev).ok,true);
    assert.equal(s.post('deleteRifa',{rifaId:item.p.rifaId},dev).ok,true);
    assert.equal(s.post('ownerStatus',{},item.token).order.status,'deleted');
    assert.equal(s.post('approveCard',{rifaId:item.p.rifaId},dev).code,'NOT_FOUND');
    assert.equal(s.post('orderCard',item.p).order.status,'deleted');
    assert.deepEqual(s.get('getSnapshot',{rifaId:item.p.rifaId}).rifas,[]);
  }
  assert.equal(s.get('getSnapshot',{rifaId:'test'}).rifas.length,1);
});
test('developer deletion removes one active raffle and its winner while preserving other raffles',()=>{
  const s=server(),dev=s.login(),a=ordered(s),b=ordered(s);
  for(const item of [a,b]) {
    s.post('approveCard',{rifaId:item.p.rifaId},dev);
    s.post('reserveNumber',reservation({rifaId:item.p.rifaId}));
    assert.equal(s.post('publishWinner',reservation({rifaId:item.p.rifaId,revision:0}),item.token).ok,true);
  }
  assert.equal(s.post('deleteRifa',{rifaId:a.p.rifaId},dev).ok,true);
  assert.equal(s.post('adminSnapshot',{rifaId:a.p.rifaId},a.token).code,'PAYMENT_PENDING');
  assert.deepEqual(s.get('getSnapshot',{rifaId:a.p.rifaId}).rifas,[]);
  assert.equal(s.get('getSnapshot',{rifaId:b.p.rifaId}).rifas[0].winner.nombre,'Ana Pérez');
  assert.equal(s.get('getSnapshot',{rifaId:b.p.rifaId}).data.length,1);
});
test('creating a card charges a fixed server amount and stores only a hashed owner key',()=>{
  const s=server(),p=card({amount:0,status:'active'}),a=ordered(s,p);
  assert.equal(a.session.order.amount,3500);assert.equal(a.session.order.status,'pending');
  assert(!JSON.stringify(s.sheets.get('RifaPro_Cartones').data).includes(p.ownerKey));
  assert.equal(s.post('orderCard',p).ok,true);assert.equal(s.sheets.get('RifaPro_Cartones').data.length,2);
  assert.equal(s.post('orderCard',{...p,ownerKey:crypto.randomUUID()+crypto.randomUUID()}).code,'CONFLICT');
});
test('unpaid cards cannot be viewed publicly, edited, reserved or activated by their owner',()=>{
  const s=server(),{p,token}=ordered(s);
  assert.equal(s.get('getSnapshot',{rifaId:p.rifaId}).rifas.length,0);
  assert.equal(s.get('getRifaNumbers',{rifaId:p.rifaId}).code,'NOT_FOUND');
  assert.equal(s.post('reserveNumber',reservation({rifaId:p.rifaId})).code,'NOT_FOUND');
  assert.equal(s.post('adminSnapshot',{rifaId:p.rifaId},token).code,'PAYMENT_PENDING');
  assert.equal(s.post('updateNumber',reservation({rifaId:p.rifaId}),token).code,'PAYMENT_PENDING');
  assert.equal(s.post('approveCard',{rifaId:p.rifaId},token).code,'FORBIDDEN');
  assert.equal(s.post('createRaffle',p,token).code,'FORBIDDEN');
});
test('reporting payment does not activate a card; platform approval is required and idempotent',()=>{
  const s=server(),{p,token}=ordered(s),admin=s.login();
  assert.equal(s.post('reportPayment',{},token).order.status,'reviewing');
  assert.equal(s.get('getSnapshot',{rifaId:p.rifaId}).rifas.length,0);
  assert.equal(s.post('listOrders',{},token).code,'FORBIDDEN');
  const queue=s.post('listOrders',{},admin);assert.equal(queue.orders[0].rifaId,p.rifaId);assert(!JSON.stringify(queue).includes('ownerHash'));
  assert.equal(s.post('approveCard',{rifaId:p.rifaId},admin).order.status,'active');
  assert.equal(s.post('approveCard',{rifaId:p.rifaId},admin).ok,true);
  assert.equal(s.get('getSnapshot',{rifaId:p.rifaId}).rifas[0].rifaId,p.rifaId);
  assert.equal(s.post('adminSnapshot',{rifaId:p.rifaId},token).rifas.length,1);
});
test('owners and shared links are isolated from every other card and legacy raffle',()=>{
  const s=server(),a=ordered(s),b=ordered(s),admin=s.login();
  for(const x of [a,b])assert.equal(s.post('approveCard',{rifaId:x.p.rifaId},admin).ok,true);
  const snapshot=s.post('adminSnapshot',{},a.token);assert.deepEqual(snapshot.rifas.map(r=>r.rifaId),[a.p.rifaId]);
  for(const action of ['adminSnapshot','updateNumber','releaseNumber','updatePrize','deleteRifa'])assert.equal(s.post(action,reservation({rifaId:b.p.rifaId}),a.token).code,'FORBIDDEN',action);
  const batch=s.post('batchNumbers',{changes:[{action:'updateNumber',payload:reservation({rifaId:b.p.rifaId,estado:'pagado',paymentMethod:'efectivo',revision:0})}]},a.token);assert.equal(batch.results[0].code,'FORBIDDEN');
  assert.deepEqual(s.get('getRifas',{rifaId:'test'}).data.map(r=>r.rifaId),['test']);
  assert.deepEqual(s.get('getSnapshot',{rifaId:a.p.rifaId}).rifas.map(r=>r.rifaId),[a.p.rifaId]);
  assert.equal(s.post('ownerLogin',{rifaId:b.p.rifaId,ownerKey:a.p.ownerKey}).code,'INVALID_KEY');
  assert.equal(s.post('updateNumber',reservation({rifaId:a.p.rifaId,estado:'pagado',paymentMethod:'efectivo',revision:0}),a.token).ok,true);
  assert.equal(s.post('adminSnapshot',{},a.token).summary.recaudado,5000);
});
test('public card data never includes owner credentials, platform payment contact or financial totals',()=>{
  const s=server(),{p,token}=ordered(s);s.post('approveCard',{rifaId:p.rifaId},s.login());s.post('reserveNumber',reservation({rifaId:p.rifaId}));
  const json=JSON.stringify(s.get('getSnapshot',{rifaId:p.rifaId}));
  for(const secret of [p.ownerKey,p.contact,'ownerHash','summary','clienteTelefono'])assert(!json.includes(secret),secret);
  s.post('logout',{sessionToken:token});assert.equal(s.post('ownerStatus',{},token).code,'UNAUTHORIZED');
});
test('expiry closes at the end of the second day after the draw in Colombia time',()=>{
  const s=server();assert.equal(s.ctx.expiry_('2026-10-31'),'2026-11-03T05:00:00.000Z');
  assert.equal(s.ctx.expiry_('invalid'),null);
  const real=Date.now;try{Date.now=()=>Date.parse('2026-11-03T04:59:59.999Z');assert.equal(s.post('reserveNumber',reservation()).ok,true);
    Date.now=()=>Date.parse('2026-11-03T05:00:00.000Z');assert.equal(s.post('reserveNumber',reservation({numeroId:'08'})).code,'EXPIRED');
    const admin=s.login();assert.equal(s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'efectivo',revision:1}),admin).code,'EXPIRED');
    assert.equal(s.post('updatePrize',{rifaId:'test',premio:'Otro',revision:0,requestId:crypto.randomUUID()},admin).code,'EXPIRED');
    assert.equal(s.get('getSnapshot',{rifaId:'test'}).rifas[0].expired,true);assert.equal(s.post('adminSnapshot',{rifaId:'test'},admin).data.length,1);
  }finally{Date.now=real;}
});
test('invalid or expired draft dates are rejected before requesting payment',()=>{
  const s=server();assert.equal(s.post('orderCard',card({fecha:'2026-02-31'})).ok,false);assert.equal(s.post('orderCard',card({fecha:'2000-01-01'})).code,'EXPIRED');assert.equal(s.sheets.get('RifaPro_Cartones').data.length,1);
});
test('deleted cards cannot be restored by replaying payment approval or an old creation',()=>{
  const s=server(),{p,token}=ordered(s),admin=s.login();s.post('approveCard',{rifaId:p.rifaId},admin);
  assert.equal(s.post('deleteRifa',{rifaId:p.rifaId},token).ok,true);
  assert.equal(s.post('approveCard',{rifaId:p.rifaId},admin).code,'NOT_FOUND');
  assert.equal(s.post('orderCard',p).order.status,'deleted');assert.equal(s.get('getSnapshot',{rifaId:p.rifaId}).rifas.length,0);
});
test('raffle password cannot log into the developer panel or approve cards',()=>{
  const s=server(),owner=s.post('login',{pin:'4826',rifaId:'test'});
  assert.equal(owner.role,'owner');assert.equal(owner.rifaId,'test');assert.equal(owner.order.legacy,true);
  assert.equal(s.post('developerLogin',{password:'4826'}).ok,false);
  for(const action of ['developerDashboard','listOrders','approveCard','createRaffle'])assert.equal(s.post(action,{rifaId:'test'},owner.token).code,'FORBIDDEN',action);
  assert.equal(s.post('adminSnapshot',{rifaId:'test'},owner.token).rifas.length,1);
  const independent=ordered(s);s.post('approveCard',{rifaId:independent.p.rifaId},s.login());
  assert.equal(s.post('adminSnapshot',{rifaId:independent.p.rifaId},owner.token).code,'FORBIDDEN');
});
test('developer sees every card including unpaid, active and legacy; public sees only its shared card',()=>{
  const s=server(),a=ordered(s),b=ordered(s),developer=s.login();s.post('approveCard',{rifaId:a.p.rifaId},developer);
  assert.equal(s.post('developerDashboard',{},b.token).code,'FORBIDDEN');
  const dashboard=s.post('developerDashboard',{},developer);assert.equal(dashboard.cards.length,3);
  assert.equal(dashboard.cards.find(c=>c.rifaId===b.p.rifaId).status,'pending');
  assert.equal(dashboard.cards.find(c=>c.rifaId===a.p.rifaId).status,'active');
  assert.equal(dashboard.cards.find(c=>c.rifaId==='test').legacy,true);
  assert.equal(s.get('getSnapshot').rifas.length,0);assert.equal(s.get('getRifas').data.length,0);
  assert.deepEqual(s.get('getSnapshot',{rifaId:'test'}).rifas.map(r=>r.rifaId),['test']);
});
test('legacy password is bound to exactly one raffle when multiple legacy cards exist',()=>{
  const s=server(),p={rifaId:'second',titulo:'Otra',cantidad:100,precio:20000,premio:'Premio',loteria:'Boyacá',fecha:'2099-10-31',hora:'17:00',responsable:'Otro'};
  s.post('createRaffle',p,s.login());assert.equal(s.post('login',{pin:'0000'}).ok,false);
  s.properties.set('LEGACY_RAFFLE_ID','test');assert.equal(s.post('login',{pin:'4826',rifaId:'second'}).code,'PIN_NOT_SET');
  const owner=s.post('login',{pin:'4826',rifaId:'test'});assert.equal(owner.role,'owner');
  assert.equal(s.post('adminSnapshot',{rifaId:'second'},owner.token).code,'FORBIDDEN');
});
test('old global sessions are revoked and developer requires a separately configured password',()=>{
  const s=server(),token=crypto.randomUUID()+crypto.randomUUID();
  for(const raw of [String(Date.now()+60000),JSON.stringify({role:'platform',expiresAt:Date.now()+60000})]){s.cache.set('session:'+token,raw);assert.equal(s.post('developerDashboard',{},token).code,'UNAUTHORIZED');}
  s.properties.delete('DEVELOPER_PIN');assert.equal(s.post('developerLogin',{password:'4826'}).ok,false);
  s.properties.set('DEVELOPER_PIN','7462');const session=s.post('developerLogin',{password:'7462'});assert.equal(session.role,'developer');
  assert.equal(s.post('developerDashboard',{},session.token).ok,true);
});
test('developer and raffle login attempt limits are separate',()=>{
  const s=server();for(let i=0;i<8;i++)s.post('login',{pin:'wrong',rifaId:'test'});
  assert.equal(s.post('developerLogin',{password:'7462'}).role,'developer');
  for(let i=0;i<8;i++)s.post('developerLogin',{password:'wrong'});
  assert.equal(s.post('developerLogin',{password:'7462'}).code,'RATE_LIMIT');
});
test('each approved card logs in with its own PIN and stays limited to that raffle',()=>{
  const s=server(),a=ordered(s),b=ordered(s,card({adminPin:'7931'})),dev=s.login();
  assert.equal(s.post('login',{rifaId:a.p.rifaId,pin:a.p.adminPin}).code,'PAYMENT_PENDING');
  for(const x of [a,b])s.post('approveCard',{rifaId:x.p.rifaId},dev);
  const session=s.post('login',{rifaId:a.p.rifaId,pin:a.p.adminPin});assert.equal(session.role,'owner');assert.equal(session.rifaId,a.p.rifaId);
  assert.equal(s.post('login',{rifaId:b.p.rifaId,pin:a.p.adminPin}).code,'INVALID_PIN');
  assert.equal(s.post('adminSnapshot',{rifaId:b.p.rifaId},session.token).code,'FORBIDDEN');
  assert.equal(s.post('approveCard',{rifaId:b.p.rifaId},session.token).code,'FORBIDDEN');
  assert.equal(s.post('login',{pin:a.p.adminPin}).rifaId,a.p.rifaId);
});
test('PINs have separate salts, stay out of public responses, and do not create per-card script properties',()=>{
  const s=server(),before=[...s.properties.keys()].sort(),a=ordered(s),b=ordered(s),dev=s.login();
  const access=s.sheets.get('RifaPro_Accesos').data;
  const first=access.find(r=>r[0]===a.p.rifaId),second=access.find(r=>r[0]===b.p.rifaId);
  assert.notEqual(first[1],second[1]);assert.notEqual(first[2],second[2]);assert.equal(first[2].length,64);
  assert(!JSON.stringify(s.sheets.get('RifaPro_Cartones').data).includes('adminPin'));
  assert(!JSON.stringify(access).includes('4826'));assert.deepEqual([...s.properties.keys()].sort(),before);
  s.post('approveCard',{rifaId:a.p.rifaId},dev);const publicJson=JSON.stringify(s.get('getSnapshot',{rifaId:a.p.rifaId}));
  for(const field of ['pinHash','pinSalt','hasPin','adminPin',first[1],first[2]])assert(!publicJson.includes(field),field);
});
test('PIN lockout persists per card, does not lock other organizers, and resets after five minutes',()=>{
  const s=server(),a=ordered(s),b=ordered(s),dev=s.login();for(const x of [a,b])s.post('approveCard',{rifaId:x.p.rifaId},dev);
  for(let i=0;i<8;i++)assert.equal(s.post('login',{rifaId:a.p.rifaId,pin:'wrong'}).code,'INVALID_PIN');
  assert.equal(s.post('login',{rifaId:a.p.rifaId,pin:a.p.adminPin}).code,'RATE_LIMIT');
  assert.equal(s.post('login',{rifaId:b.p.rifaId,pin:b.p.adminPin}).role,'owner');
  s.cache.clear();assert.equal(s.post('login',{rifaId:a.p.rifaId,pin:a.p.adminPin}).code,'RATE_LIMIT');
  const real=Date.now;try{const now=real();Date.now=()=>now+300001;assert.equal(s.post('login',{rifaId:a.p.rifaId,pin:a.p.adminPin}).role,'owner');}finally{Date.now=real;}
});
test('existing owners can set their missing PIN without affecting another card or replacing a saved PIN',()=>{
  const s=server(),a=ordered(s),b=ordered(s),dev=s.login();s.post('approveCard',{rifaId:a.p.rifaId},dev);
  const access=s.sheets.get('RifaPro_Accesos');access.data=access.data.filter(r=>r[0]!==a.p.rifaId);
  assert.equal(s.post('setOwnerPin',{rifaId:a.p.rifaId,pin:'7931'},b.token).code,'FORBIDDEN');
  assert.equal(s.post('setOwnerPin',{rifaId:a.p.rifaId,pin:'123'},a.token).code,'INVALID_PIN_FORMAT');
  assert.equal(s.post('setOwnerPin',{rifaId:a.p.rifaId,pin:'7931'},a.token).order.hasPin,true);
  assert.equal(s.post('setOwnerPin',{rifaId:a.p.rifaId,pin:b.p.adminPin},a.token).code,'PIN_IN_USE');
  assert.equal(s.post('orderCard',a.p).ok,true);assert.equal(s.post('login',{rifaId:a.p.rifaId,pin:'7931'}).role,'owner');
  assert.equal(s.post('login',{rifaId:a.p.rifaId,pin:a.p.adminPin}).code,'INVALID_PIN');
});
test('legacy PIN import is idempotent and no longer depends on ADMIN_PIN during login',()=>{
  const s=server();const original=JSON.stringify(s.sheets.get('RifaPro_Accesos').data);s.ctx.prepararSistema();assert.equal(JSON.stringify(s.sheets.get('RifaPro_Accesos').data),original);
  s.properties.delete('ADMIN_PIN');assert.equal(s.post('login',{rifaId:'test',pin:'4826'}).role,'owner');
});
test('PIN-only login restores saved reservations, payments and prize for exactly one card',()=>{
  const s=server(),a=ordered(s),dev=s.login();s.post('approveCard',{rifaId:a.p.rifaId},dev);
  const p=reservation({rifaId:a.p.rifaId,estado:'pagado',paymentMethod:'efectivo',revision:0});
  assert.equal(s.post('updateNumber',p,a.token).ok,true);
  s.post('logout',{sessionToken:a.token});
  const session=s.post('login',{pin:a.p.adminPin});assert.equal(session.rifaId,a.p.rifaId);
  const snap=s.post('adminSnapshot',{},session.token);assert.equal(snap.rifas.length,1);
  assert.equal(snap.rifas[0].premio,a.p.premio);assert.equal(snap.data[0].clienteNombre,p.clienteNombre);assert.equal(snap.summary.efectivo,a.p.precio);
});
test('duplicate and developer PINs reject creation before inserting an order',()=>{
  const s=server(),a=ordered(s),before=s.sheets.get('RifaPro_Cartones').data.length;
  for(const pin of [a.p.adminPin,'7462'])assert.equal(s.post('orderCard',card({adminPin:pin})).code,'PIN_IN_USE');
  assert.equal(s.sheets.get('RifaPro_Cartones').data.length,before);
});
test('five-column access migration supports PIN-only login without replacing the salted credential',()=>{
  const s=server(),access=s.sheets.get('RifaPro_Accesos');access.data.forEach(row=>row.pop());
  const hash=access.data[1][2];s.ctx.prepararSistema();
  assert.equal(access.data[0][5],'pinLookup');assert.equal(s.post('login',{pin:'4826'}).rifaId,'test');
  assert.equal(access.data[1][2],hash);assert.equal(access.data[1][5].length,64);
});
test('ambiguous legacy PINs never select an arbitrary card; developer can resolve the conflict',()=>{
  const s=server(),a=ordered(s),dev=s.login();s.post('approveCard',{rifaId:a.p.rifaId},dev);
  const rows=s.sheets.get('RifaPro_Accesos').data,legacy=rows.find(r=>r[0]==='test'),other=rows.find(r=>r[0]===a.p.rifaId);
  other[1]=legacy[1];other[2]=legacy[2];other[5]='';legacy[5]='';
  assert.equal(s.post('login',{pin:'4826'}).code,'PIN_CONFLICT');
  assert(s.post('developerDashboard',{},dev).cards.filter(c=>c.pinConflict).length===2);
  assert.equal(s.post('setOwnerPin',{rifaId:a.p.rifaId,pin:'7931'},dev).ok,true);
  assert.equal(s.post('login',{pin:'7931'}).rifaId,a.p.rifaId);assert.equal(s.post('login',{pin:'4826'}).rifaId,'test');
});
test('global PIN miss throttling survives cache loss and a successful known login',()=>{
  const s=server();for(let i=0;i<7;i++)assert.equal(s.post('login',{pin:'9999'}).code,'INVALID_PIN');
  assert.equal(s.post('login',{pin:'4826'}).role,'owner');assert.equal(s.post('login',{pin:'9999'}).code,'INVALID_PIN');
  s.cache.clear();assert.equal(s.post('login',{pin:'4826'}).code,'RATE_LIMIT');assert.equal(s.post('developerLogin',{password:'7462'}).role,'developer');
});
test('owner PIN change preserves numbers and rejects changes to another owner',()=>{
  const s=server(),a=ordered(s),b=ordered(s),dev=s.login();s.post('approveCard',{rifaId:a.p.rifaId},dev);
  s.post('updateNumber',reservation({rifaId:a.p.rifaId,revision:0,estado:'pagado',paymentMethod:'transferencia'}),a.token);
  assert.equal(s.post('setOwnerPin',{rifaId:b.p.rifaId,pin:'7931'},a.token).code,'FORBIDDEN');
  assert.equal(s.post('setOwnerPin',{rifaId:a.p.rifaId,pin:'7931'},a.token).ok,true);
  assert.equal(s.post('login',{pin:a.p.adminPin}).code,'INVALID_PIN');
  const owner=s.post('login',{pin:'7931'});assert.equal(s.post('adminSnapshot',{},owner.token).summary.transferencia,a.p.precio);
});
test('short references remain unique and stable through retries, approval and deletion',()=>{
  const s=server(),a=ordered(s),b=ordered(s),dev=s.login();
  const first=s.post('ownerStatus',{},a.token).order.reference,second=s.post('ownerStatus',{},b.token).order.reference;
  assert.match(first,/^RP-\d{5}$/);assert.notEqual(first,second);
  assert.equal(s.post('orderCard',a.p).order.reference,first);
  assert.equal(s.post('approveCard',{rifaId:a.p.rifaId},dev).order.reference,first);
  assert.equal(s.post('deleteRifa',{rifaId:a.p.rifaId},dev).ok,true);
  assert.equal(s.post('ownerStatus',{},b.token).order.reference,second);
  assert.equal(s.post('listOrders',{},dev).orders.find(o=>o.rifaId===a.p.rifaId).reference,first);
});
test('public colors follow confirmed cash, Nequi, reserved and released states without exposing contacts',()=>{
  const s=server(),token=s.login();
  for(const [revision,estado,paymentMethod] of [[0,'pagado','transferencia'],[1,'pagado','efectivo'],[2,'apartado','']]){
    assert.equal(s.post('updateNumber',reservation({revision,estado,paymentMethod}),token).ok,true);
    const snap=s.get('getSnapshot',{rifaId:'test'}),record=snap.data[0];assert.equal(record.estado,estado);assert.equal(record.paymentMethod,paymentMethod);
    assert(!JSON.stringify(record).includes('3001234567'));assert(!('summary' in snap));
  }
  s.post('releaseNumber',reservation({revision:3}),token);const record=s.get('getRifaNumbers',{rifaId:'test'}).data[0];
  assert.equal(record.estado,'libre');assert.equal(record.paymentMethod,'');assert.equal(record.clienteNombre,'');
});
test('prize creation and edits preserve full peso amounts and non-cash prizes',()=>{
  const s=server(),p=card({premio:'500000'}),dev=s.login();assert.equal(s.post('orderCard',p).ok,true);
  assert.equal(s.post('approveCard',{rifaId:p.rifaId},dev).ok,true);
  assert.equal(s.get('getSnapshot',{rifaId:p.rifaId}).rifas[0].premio,'$500.000');
  for(const [revision,premio,expected] of [[0,'800.000','$800.000'],[1,'$1.200.000','$1.200.000'],[2,'Una bicicleta','Una bicicleta']]){
    const result=s.post('updatePrize',{rifaId:p.rifaId,revision,premio,requestId:crypto.randomUUID()},dev);
    assert.equal(result.rifa.premio,expected);assert.equal(s.get('getSnapshot',{rifaId:p.rifaId}).rifas[0].premio,expected);
  }
});
