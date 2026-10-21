const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const crypto = require('node:crypto');
const code = fs.readFileSync('backend/Code.gs', 'utf8');

function server(legacyRows) {
  const sheets = new Map(), properties = new Map([['SPREADSHEET_ID','test'],['ADMIN_PIN','4826']]), cache = new Map();
  let locked = false;
  class Sheet {
    constructor() { this.data = []; }
    appendRow(row) { this.data.push([...row]); }
    setFrozenRows() {}
    getLastRow() { return this.data.length; }
    deleteRow(row) { assert(locked); this.data.splice(row - 1, 1); }
    getRange(start, column, count, width) {
      return { getValues: () => Array.from({length:count},(_,i) => Array.from({length:width},(_,j) => this.data[start-1+i]?.[column-1+j] ?? '')),
        setNumberFormat() {}, setValues: (values) => { assert(locked, 'writes require the server lock'); values.forEach((r,i) => { this.data[start-1+i] ||= []; r.forEach((v,j) => {this.data[start-1+i][column-1+j]=v;}); }); } };
    }
  }
  const db = { getId: () => 'test', getSheetByName: (name) => sheets.get(name), insertSheet: (name) => { const s = new Sheet(); sheets.set(name,s); return s; } };
  if (legacyRows) { const sheet = new Sheet(); sheet.data = legacyRows.map(row=>[...row]); sheets.set('rifas',sheet); }
  const ctx = vm.createContext({ console, Date, JSON, Math, String, Number, Error,
    SpreadsheetApp: { openById: () => db, getActiveSpreadsheet: () => db, flush() {} },
    LockService: { getScriptLock: () => ({ tryLock: () => { assert(!locked); locked = true; return true; }, releaseLock: () => { locked = false; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (key) => properties.get(key), setProperty: (k,v) => properties.set(k,v), deleteProperty: (key) => properties.delete(key) }) },
    CacheService: { getScriptCache: () => ({ get: (key) => cache.get(key), put: (k,v) => cache.set(k,v), remove: (key) => cache.delete(key) }) },
    Utilities: { getUuid: () => crypto.randomUUID(), formatDate: (value,zone,format) => { const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:zone,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).formatToParts(value).map(p=>[p.type,p.value])); return format==='HH:mm'?`${parts.hour}:${parts.minute}`:`${parts.year}-${parts.month}-${parts.day}`; } },
    ContentService: { MimeType: { JSON:'json', JAVASCRIPT:'js' }, createTextOutput: (value) => ({ value, setMimeType() { return this; } }) }
  });
  vm.runInContext(code,ctx);
  ctx.prepararSistema();
  if (!legacyRows) ctx.createRaffle_({ rifaId:'test',titulo:'Rifa',cantidad:100,precio:20000,premio:'$500.000',loteria:'Boyacá',fecha:'2026-10-31',hora:'17:00',responsable:'Organizadora' });
  const post = (action,payload = {}, token) => JSON.parse(ctx.doPost({ postData:{contents:JSON.stringify({action,payload,token})} }).value);
  const get = (action,params = {}) => JSON.parse(ctx.doGet({ parameter:{action,...params} }).value);
  const login = () => post('login',{pin:'4826'}).token;
  return {ctx,post,get,login,cache,properties,sheets};
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
  for (const action of ['adminSnapshot','updateNumber','releaseNumber','deleteRifa','createRaffle']) assert.equal(s.post(action,reservation()).code,'UNAUTHORIZED');
  assert.equal(s.get('verifyPin').ok,false); assert.equal(s.get('upsertNumber').ok,false);
});
test('public snapshot hides phone, actual paid state, payment method and revenue', () => {
  const s = server(), token = s.login(); s.post('updateNumber',reservation({estado:'pagado',paymentMethod:'efectivo',revision:0}),token);
  const pub = s.get('getRifaNumbers',{rifaId:'test'}); assert.equal(pub.data[0].estado,'apartado');
  const json = JSON.stringify(pub); for (const secret of ['3001234567','paymentMethod','efectivo','recaudado','revision','requestId']) assert(!json.includes(secret),secret);
});
test('Configured PIN works; incorrect PIN fails; logout revokes access', () => {
  const s = server(); assert.equal(s.post('login',{pin:'9999'}).ok,false); const token = s.login(); assert(token);
  assert.equal(s.post('adminSnapshot',{rifaId:'test'},token).ok,true); s.post('logout',{sessionToken:token});
  assert.equal(s.post('adminSnapshot',{rifaId:'test'},token).code,'UNAUTHORIZED');
});
test('repeated wrong PINs are throttled across client identities', () => {
  const s = server(); for (let i=0;i<8;i++) s.post('login',{pin:'9999',clientId:crypto.randomUUID()});
  assert.equal(s.post('login',{pin:'4826'}).code,'RATE_LIMIT');
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
  assert.equal(s.get('getRifas').data.length,1); s.post('reserveNumber',reservation());
  assert.equal(s.post('deleteRifa',{rifaId:'test'},token).ok,true); assert.equal(s.get('getRifas').data.length,0);
  assert.equal(s.post('reserveNumber',reservation()).code,'NOT_FOUND');
});
test('JSONP rejects executable callback names and exposes read actions only', () => {
  const s = server(); const out = s.ctx.doGet({parameter:{action:'getRifas',callback:'alert(1)//'}}).value;
  assert.equal(JSON.parse(out).ok,false); assert(!out.includes('alert(1)'));
  const valid = s.ctx.doGet({parameter:{action:'getRifas',callback:'rifa_valid123'}}).value; assert(valid.startsWith('rifa_valid123('));
});
test('setup is repeatable and does not delete existing data', () => {
  const s = server(); s.post('reserveNumber',reservation()); s.ctx.prepararSistema();
  assert.equal(s.get('getRifas').data.length,1); assert.equal(s.get('getRifaNumbers',{rifaId:'test'}).data.length,1);
});

const legacyHeaders=['rifaId','numeroId','estado','clienteNombre','clienteTelefono','updatedAt','titulo','precio','premio','loteria','fecha','hora','responsable'];
const legacyMeta=['old','__META__','meta','100','','2026-10-01T10:00:00Z','Rifa anterior',20000,'$500.000','Boyacá','2026-10-31','17:00','Organizadora'];
const legacyPaid=['old','07','pagado','Ana anterior','3007654321','2026-10-02T10:00:00Z','Rifa anterior',20000,'$500.000','Boyacá','2026-10-31','17:00','Organizadora'];
test('original 13-column spreadsheet and __META__ rows remain readable without migration', () => {
  const s=server([legacyHeaders,legacyMeta,legacyPaid]);
  const r=s.get('getRifas');assert.equal(r.data[0].titulo,'Rifa anterior');assert.equal(r.data[0].cantidad,100);
  const snap=s.post('adminSnapshot',{rifaId:'old'},s.login());assert.equal(snap.data[0].clienteNombre,'Ana anterior');assert.equal(snap.data[0].clienteTelefono,'3007654321');
  assert.equal(snap.summary.recaudado,20000);assert.equal(snap.summary.sinClasificar,20000);
  assert.deepEqual(s.sheets.get('rifas').data[1],legacyMeta);assert.deepEqual(s.sheets.get('rifas').data[2],legacyPaid);
  assert.deepEqual(s.sheets.get('rifas').data[0].slice(0,13),legacyHeaders);
});
test('an existing paid number cannot be reserved and its private payment stays private', () => {
  const s=server([legacyHeaders,legacyMeta,legacyPaid]);
  const result=s.post('reserveNumber',reservation({rifaId:'old'}));assert.equal(result.code,'CONFLICT');assert.equal(result.record.clienteNombre,'Ana anterior');
  const pub=s.get('getRifaNumbers',{rifaId:'old'});assert.equal(pub.data[0].estado,'apartado');assert(!JSON.stringify(pub).includes('3007654321'));
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
  const s=server([legacyHeaders,meta]);const r=s.get('getRifas').data[0];assert.equal(r.fecha,'2026-10-31');assert.equal(r.hora,'17:00');
});
test('unrelated raffles and their rows survive deletion', () => {
  const other=[...legacyMeta];other[0]='other';const s=server([legacyHeaders,legacyMeta,legacyPaid,other]);
  assert.equal(s.post('deleteRifa',{rifaId:'old'},s.login()).ok,true);assert.equal(s.get('getRifas').data[0].rifaId,'other');
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
