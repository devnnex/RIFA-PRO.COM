/* Google Apps Script V8. Compatible with the original "rifas" spreadsheet. */
const PROTOCOL = 2;
const SHEET_NAME = "rifas";
const META_NUMBER_ID = "__META__";
const HEADERS = ["rifaId", "numeroId", "estado", "clienteNombre", "clienteTelefono", "updatedAt",
  "titulo", "precio", "premio", "loteria", "fecha", "hora", "responsable",
  "paymentMethod", "revision", "requestId", "condiciones"];
const TABLES = {
  entries: { name: SHEET_NAME, headers: HEADERS },
  requests: { name: "RifaPro_v2_Solicitudes", headers: ["requestId", "action", "rifaId", "numeroId", "result"] }
};
const SESSION_SECONDS = 1800;
const CAPABILITIES = { snapshot: true, editPrize: true, batchNumbers: true, writeSnapshot: true };
let requestContext = null;
let stagedBatch = null;
function resetContext_() { requestContext = { db: null, sheets: {}, rows: {} }; }

function doGet(event) {
  resetContext_();
  const p = event.parameter || {};
  // JSONP is read-only and contains public information only.
  let result;
  try {
    if (p.action === "health") result = { ok: true, protocol: PROTOCOL, capabilities: CAPABILITIES, now: new Date().toISOString(), sheet: SHEET_NAME };
    else if (p.action === "getRifas") result = { ok: true, protocol: PROTOCOL, capabilities: CAPABILITIES, data: publicCatalog_() };
    else if (p.action === "getSnapshot") {
      const rifas = publicCatalog_();
      const rifa = rifas.find((r) => r.rifaId === p.rifaId) || rifas[0];
      result = { ok: true, protocol: PROTOCOL, capabilities: CAPABILITIES, rifas,
        rifaId: rifa ? rifa.rifaId : null, data: rifa ? publicSnapshot_(rifa.rifaId) : [] };
    }
    else if (p.action === "getRifaNumbers") {
      const rifa = raffle_(p.rifaId);
      result = { ok: true, protocol: PROTOCOL, data: publicSnapshot_(rifa.rifaId) };
    } else result = { ok: false, error: "Acción pública no disponible. Actualiza la aplicación." };
  } catch (error) { result = failure_(error); }
  const callback = p.callback;
  if (callback && !/^rifa_[a-zA-Z0-9_]{1,80}$/.test(callback)) return output_({ ok: false, error: "Callback inválido." });
  return output_(result, callback);
}
function doPost(event) {
  resetContext_();
  try {
    const request = JSON.parse(event.postData.contents);
    const payload = request.payload || {};
    if (request.action === "login") return output_(login_(payload.pin));
    if (request.action === "logout") {
      const token = String(payload.sessionToken || "");
      if (/^[a-f0-9-]{60,100}$/.test(token)) CacheService.getScriptCache().remove(`session:${token}`);
      return output_({ ok: true });
    }
    if (request.action === "reserveNumber") return output_(reserve_(payload));
    requireAdmin_(request.token);
    if (request.action === "adminSnapshot") {
      const rifas = readRaffles_();
      const rifa = rifas.find((r) => r.rifaId === payload.rifaId) || rifas[0];
      // Read once: summary and displayed numbers refer to the same snapshot.
      const records = rifa ? numberRows_(rifa.rifaId).map((n) => n.record) : [];
      return output_({ ok: true, protocol: PROTOCOL, capabilities: CAPABILITIES, rifas,
        rifaId: rifa ? rifa.rifaId : null, data: records, summary: rifa ? summary_(rifa, records) : null });
    }
    if (request.action === "createRaffle") return output_(createRaffle_(payload));
    if (request.action === "updatePrize") return output_(updatePrize_(payload));
    if (request.action === "batchNumbers") return output_(batchNumbers_(payload));
    if (request.action === "updateNumber" || request.action === "releaseNumber") return output_(editNumber_(request.action, payload));
    if (request.action === "deleteRifa") return output_(deleteRaffle_(payload.rifaId));
    return output_({ ok: false, error: "Acción no disponible." });
  } catch (error) { return output_(failure_(error)); }
}
function output_(result, callback) {
  const json = JSON.stringify(result).replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029");
  return ContentService.createTextOutput(callback ? `${callback}(${json});` : json)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
function failure_(error) {
  return { ok: false, code: error.code || (error.safe ? "INVALID" : "RETRY"), error: error.safe ? error.message : "No se pudo confirmar la solicitud. Intenta de nuevo con la misma operación." };
}
function fail_(message, code) { const error = new Error(message); error.safe = true; error.code = code || "INVALID"; throw error; }
function db_() {
  if (!requestContext) resetContext_();
  if (requestContext.db) return requestContext.db;
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const db = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!db) fail_("Configura SPREADSHEET_ID o ejecuta prepararSistema desde el proyecto vinculado a tu hoja.");
  requestContext.db = db;
  return db;
}
function table_(key) {
  if (!requestContext) resetContext_();
  if (requestContext.sheets[key]) return requestContext.sheets[key];
  const sheet = db_().getSheetByName(TABLES[key].name);
  if (!sheet) fail_("El organizador debe ejecutar prepararSistema antes de publicar.");
  requestContext.sheets[key] = sheet;
  return sheet;
}
function rows_(key) {
  if (!requestContext) resetContext_();
  if (requestContext.rows[key]) return requestContext.rows[key];
  const sheet = table_(key);
  const count = sheet.getLastRow() - 1;
  requestContext.rows[key] = count > 0 ? sheet.getRange(2, 1, count, TABLES[key].headers.length).getValues() : [];
  return requestContext.rows[key];
}
function text_(value, max) { return String(value == null ? "" : value).trim().slice(0, max); }
function safeCell_(value) { const s = String(value == null ? "" : value); return /^[=+@\-]/.test(s) ? "'" + s : s; }
function unsheet_(value) { return String(value == null ? "" : value).replace(/^'(?=[=+@\-])/, ""); }
function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) fail_("Hay varias reservas en curso. Intenta de nuevo en unos segundos.", "BUSY");
  resetContext_();
  try { const result = fn(); SpreadsheetApp.flush(); return result; }
  finally { lock.releaseLock(); }
}
function isoDate_(value) { return value instanceof Date ? value.toISOString() : String(value || ""); }
function legacyField_(value, format) {
  if (value instanceof Date) return Utilities.formatDate(value, "America/Bogota", format);
  return unsheet_(value);
}
function readRaffles_() {
  const groups = new Map();
  rows_("entries").forEach((row) => {
    const id = String(row[0] || ""); if (!id) return;
    if (!groups.has(id)) groups.set(id, []);
    groups.get(id).push(row);
  });
  return Array.from(groups, ([id, records]) => {
    // Merge nonempty metadata in timestamp order, as the original server did.
    const sorted = records.slice().sort((a, b) => (Date.parse(isoDate_(a[5])) || 0) - (Date.parse(isoDate_(b[5])) || 0));
    const r = { rifaId: id, titulo: "", precio: 0, premio: "", loteria: "", fecha: "", hora: "", responsable: "", condiciones: "", cantidad: 0, revision: 0 };
    let maxNumber = -1, metaQuantity = 0;
    sorted.forEach((row) => {
      [["titulo",6],["premio",8],["loteria",9],["responsable",12],["condiciones",16]].forEach(([field,col]) => { if (row[col]) r[field] = unsheet_(row[col]); });
      if (row[7] !== "" && row[7] != null && Number.isFinite(Number(row[7]))) r.precio = Number(row[7]);
      if (row[10]) r.fecha = legacyField_(row[10], "yyyy-MM-dd");
      if (row[11]) r.hora = legacyField_(row[11], "HH:mm");
      if (String(row[1]) === META_NUMBER_ID) { if (Number(row[3]) > 0) metaQuantity = Number(row[3]); r.revision = Number(row[14]) || 0; }
      else if (/^\d+$/.test(String(row[1]))) maxNumber = Math.max(maxNumber, Number(row[1]));
    });
    r.cantidad = Math.min(100000, Math.max(metaQuantity, maxNumber + 1, 1));
    if (!metaQuantity && maxNumber < 0) r.cantidad = 100;
    return r;
  });
}
function publicCatalog_() {
  const cache = CacheService.getScriptCache(), key = "rifapro:catalog:v3";
  const cached = cache.get(key); if (cached) return JSON.parse(cached);
  const data = readRaffles_(), json = JSON.stringify(data);
  if (json.length < 80000) cache.put(key, json, 5);
  return data;
}
function raffle_(id) {
  const r = readRaffles_().find((item) => item.rifaId === String(id));
  if (!r) fail_("Esta rifa ya no está disponible.", "NOT_FOUND");
  return r;
}
function numberRows_(id) {
  const selected = new Map();
  rows_("entries").forEach((row, i) => {
    if (String(row[0]) !== String(id) || !/^\d+$/.test(String(row[1]))) return;
    const record = { rifaId: String(row[0]), numeroId: String(row[1]), estado: String(row[2] || "libre"),
      clienteNombre: unsheet_(row[3]), clienteTelefono: unsheet_(row[4]), paymentMethod: String(row[13] || ""),
      updatedAt: isoDate_(row[5]), revision: Number(row[14]) || 0, requestId: String(row[15] || "") };
    const key = Number(record.numeroId), previous = selected.get(key);
    if (!previous || (Date.parse(record.updatedAt) || 0) >= (Date.parse(previous.record.updatedAt) || 0)) selected.set(key, { row: i + 2, record });
  });
  return Array.from(selected.values());
}
function publicRecord_(record) {
  return { numeroId: record.numeroId, estado: record.estado === "libre" ? "libre" : "apartado",
    clienteNombre: record.estado === "libre" ? "" : record.clienteNombre, updatedAt: record.updatedAt };
}
function publicSnapshot_(id) {
  const cache = CacheService.getScriptCache();
  const key = `public:${id}`;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const records = numberRows_(id).map((n) => publicRecord_(n.record));
  const json = JSON.stringify(records);
  if (json.length < 80000) cache.put(key, json, 3);
  return records;
}
function saveNumber_(record, row) {
  const r = raffle_(record.rifaId);
  const values = [record.rifaId, record.numeroId, record.estado, safeCell_(record.clienteNombre), safeCell_(record.clienteTelefono),
    record.updatedAt, safeCell_(r.titulo), r.precio, safeCell_(r.premio), safeCell_(r.loteria), r.fecha, r.hora,
    safeCell_(r.responsable), record.paymentMethod, record.revision, record.requestId, safeCell_(r.condiciones)];
  const sheet = table_("entries");
  const entries = rows_("entries");
  const targetRow = row || entries.length + 2;
  if (stagedBatch) stagedBatch.writes.set(targetRow, values);
  else { const target = sheet.getRange(targetRow, 1, 1, values.length); target.setNumberFormat("@"); target.setValues([values]); }
  entries[targetRow - 2] = values;
  // Preserve the original server's duplicate cleanup for 01 / 1, without
  // touching any other participant or raffle.
  const duplicateRows = rows_("entries").map((entry, index) => ({ entry, row: index + 2 }))
    .filter((item) => item.row !== targetRow && String(item.entry[0]) === record.rifaId && /^\d+$/.test(String(item.entry[1])) && Number(item.entry[1]) === Number(record.numeroId));
  duplicateRows.reverse().forEach((item) => { sheet.deleteRow(item.row); entries.splice(item.row - 2, 1); });
  CacheService.getScriptCache().remove(`public:${record.rifaId}`);
}
function validNumber_(payload, rifa) {
  if (!/^\d{1,6}$/.test(String(payload.numeroId))) fail_("Número inválido.");
  const number = Number(payload.numeroId);
  if (!Number.isInteger(number) || number < 0 || number >= rifa.cantidad) fail_("Ese número no pertenece a la rifa.");
  return String(number).padStart(Math.max(2, String(rifa.cantidad - 1).length), "0");
}
function validateCustomer_(p) {
  const name = text_(p.clienteNombre, 80), phone = text_(p.clienteTelefono, 30);
  if (name.length < 2 || phone.replace(/\D/g, "").length < 7) fail_("Escribe un nombre y un celular válido.");
  return { name, phone };
}
function requestId_(p) {
  if (!/^[a-f0-9-]{36}$/.test(String(p.requestId))) fail_("Identificador de solicitud inválido.");
  return p.requestId;
}
function oldRequest_(action, p) {
  const id = requestId_(p);
  const row = rows_("requests").find((r) => r[0] === id);
  if (!row) return null;
  if (row[1] !== action || row[2] !== p.rifaId || Number(row[3]) !== Number(p.numeroId)) fail_("La solicitud pertenece a otra operación.");
  return JSON.parse(row[4]);
}
function remember_(action, p, result) {
  const row = [p.requestId, action, p.rifaId, String(p.numeroId), JSON.stringify(result)];
  const requests = rows_("requests");
  if (stagedBatch) stagedBatch.requests.push(row);
  else table_("requests").appendRow(row);
  requests.push(row);
  return result;
}
function reserve_(p) {
  return locked_(() => {
    const rifa = raffle_(p.rifaId), number = validNumber_(p, rifa), customer = validateCustomer_(p);
    const previous = oldRequest_("reserveNumber", p);
    if (previous) {
      // A retry must not claim success after an administrator released/reassigned it.
      const current = numberRows_(rifa.rifaId).find((n) => Number(n.record.numeroId) === Number(number));
      if (current?.record.requestId !== p.requestId || current.record.estado === "libre")
        return { ok: false, code: "CONFLICT", record: current ? publicRecord_(current.record) : null };
      return previous;
    }
    const existing = numberRows_(rifa.rifaId).find((n) => Number(n.record.numeroId) === Number(number));
    // Recover an acknowledged write if the request ledger write was interrupted.
    if (existing?.record.requestId === p.requestId && existing.record.estado !== "libre")
      return remember_("reserveNumber", p, { ok: true, record: publicRecord_(existing.record) });
    if (existing && existing.record.estado !== "libre") return { ok: false, code: "CONFLICT", record: publicRecord_(existing.record) };
    const record = { rifaId: rifa.rifaId, numeroId: number, estado: "apartado", clienteNombre: customer.name,
      clienteTelefono: customer.phone, paymentMethod: "", updatedAt: new Date().toISOString(),
      revision: (existing?.record.revision || 0) + 1, requestId: p.requestId };
    // Read + validate + write all happen while holding the same server lock.
    saveNumber_(record, existing?.row);
    return remember_("reserveNumber", p, { ok: true, record: publicRecord_(record) });
  });
}
function editNumber_(action, p) {
  return locked_(() => editNumberLocked_(action, p));
}
function numberAck_(rifaId, record) {
  return { ok: true, rifaId, record, summary: summary_(raffle_(rifaId), numberRows_(rifaId).map((n) => n.record)) };
}
function editNumberLocked_(action, p) {
    const rifa = raffle_(p.rifaId), number = validNumber_(p, rifa);
    const previous = oldRequest_(action, p);
    const existing = numberRows_(rifa.rifaId).find((n) => Number(n.record.numeroId) === Number(number));
    if (previous || existing?.record.requestId === p.requestId) return numberAck_(rifa.rifaId, existing ? existing.record : null);
    if ((existing?.record.revision || 0) !== Number(p.revision)) fail_("Otra sesión actualizó este número. Cierra y vuelve a abrirlo antes de guardar.", "STALE");
    const releasing = action === "releaseNumber";
    const customer = releasing ? { name: "", phone: "" } : validateCustomer_(p);
    const status = releasing ? "libre" : p.estado;
    if (!["libre", "apartado", "pagado"].includes(status)) fail_("Estado inválido.");
    const method = status === "pagado" ? p.paymentMethod : "";
    if (status === "pagado" && !["transferencia", "efectivo"].includes(method)) fail_("Selecciona Nequi / transferencia o efectivo.");
    const record = { rifaId: rifa.rifaId, numeroId: number, estado: status, clienteNombre: customer.name,
      clienteTelefono: customer.phone, paymentMethod: method, updatedAt: new Date().toISOString(),
      revision: (existing?.record.revision || 0) + 1, requestId: p.requestId };
    saveNumber_(record, existing?.row);
    return remember_(action, p, numberAck_(rifa.rifaId, record));
}
function batchNumbers_(p) {
  if (!Array.isArray(p.changes) || !p.changes.length || p.changes.length > 25) fail_("Envía entre 1 y 25 cambios por lote.");
  return locked_(() => {
    const entries = rows_("entries"), seen = new Set();
    const duplicate = entries.some((row) => { if (!/^\d+$/.test(String(row[1]))) return false; const key = String(row[0]) + ":" + Number(row[1]); if (seen.has(key)) return true; seen.add(key); return false; });
    // Existing duplicate rows use the regular cleanup path. Otherwise group new
    // contiguous rows and the request ledger into range writes.
    stagedBatch = duplicate ? null : { writes: new Map(), requests: [] };
    try {
      const results = p.changes.map((change) => {
        try {
          if (!["updateNumber", "releaseNumber"].includes(change.action)) fail_("Acción de lote inválida.");
          return { requestId: change.payload?.requestId, ...editNumberLocked_(change.action, change.payload || {}) };
        } catch (error) { return { requestId: change.payload?.requestId, ...failure_(error) }; }
      });
      if (stagedBatch) {
        const writes = Array.from(stagedBatch.writes).sort((a, b) => a[0] - b[0]);
        for (let i = 0; i < writes.length;) {
          const first = writes[i][0], values = [writes[i][1]]; i++;
          while (i < writes.length && writes[i][0] === first + values.length) values.push(writes[i++][1]);
          const range = table_("entries").getRange(first, 1, values.length, HEADERS.length); range.setNumberFormat("@"); range.setValues(values);
        }
        if (stagedBatch.requests.length) table_("requests").getRange(table_("requests").getLastRow() + 1, 1, stagedBatch.requests.length, 5).setValues(stagedBatch.requests);
      }
      return { ok: true, results };
    } finally { stagedBatch = null; }
  });
}
function summary_(rifa, records) {
  const paid = records.filter((r) => r.estado === "pagado");
  const reserved = records.filter((r) => r.estado === "apartado");
  const count = (method) => paid.filter((r) => r.paymentMethod === method).length * rifa.precio;
  const transfer = count("transferencia"), cash = count("efectivo");
  return { libres: rifa.cantidad - paid.length - reserved.length, apartados: reserved.length, pagados: paid.length,
    recaudado: paid.length * rifa.precio, meta: rifa.cantidad * rifa.precio,
    transferencia: transfer, efectivo: cash, sinClasificar: paid.length * rifa.precio - transfer - cash };
}
function login_(pin) {
  return locked_(() => {
    const props = PropertiesService.getScriptProperties();
    const expected = props.getProperty("ADMIN_PIN");
    if (!expected || !/^\d{4}$/.test(expected)) fail_("El organizador debe configurar su contraseña en el servidor.");
    // Persistent global throttle: cannot be bypassed by changing browser/client ID.
    let throttle = JSON.parse(props.getProperty("LOGIN_THROTTLE") || "{}");
    if (throttle.until > Date.now()) fail_("Demasiados intentos. Espera unos minutos para volver a entrar.", "RATE_LIMIT");
    if (String(pin) !== expected) {
      if (!throttle.start || Date.now() - throttle.start > 300000) throttle = { start: Date.now(), attempts: 0 };
      throttle.attempts = (throttle.attempts || 0) + 1;
      if (throttle.attempts >= 8) throttle.until = Date.now() + 300000;
      props.setProperty("LOGIN_THROTTLE", JSON.stringify(throttle));
      fail_("Contraseña incorrecta.", "INVALID_PIN");
    }
    props.deleteProperty("LOGIN_THROTTLE");
    const token = Utilities.getUuid() + Utilities.getUuid();
    const expiresAt = Date.now() + SESSION_SECONDS * 1000;
    CacheService.getScriptCache().put(`session:${token}`, String(expiresAt), SESSION_SECONDS);
    return { ok: true, token, expiresAt };
  });
}
function requireAdmin_(token) {
  if (!/^[a-f0-9-]{60,100}$/.test(String(token || ""))) fail_("Debes ingresar como administrador.", "UNAUTHORIZED");
  const expires = Number(CacheService.getScriptCache().get(`session:${token}`));
  if (!expires || expires <= Date.now()) fail_("Tu sesión terminó. Ingresa de nuevo.", "UNAUTHORIZED");
}
function createRaffle_(p) {
  return locked_(() => {
    const id = text_(p.rifaId, 80);
    if (!/^[a-zA-Z0-9_-]{1,80}$/.test(id)) fail_("Identificador de rifa inválido.");
    const existing = readRaffles_().find((r) => r.rifaId === id);
    if (existing) return { ok: true, rifaId: id, rifa: existing };
    const quantity = Number(p.cantidad), price = Number(p.precio);
    if (!Number.isInteger(quantity) || quantity < 2 || quantity > 100000 || !Number.isSafeInteger(price) || price < 1 || price > 100000000) fail_("Revisa el precio y la cantidad de números.");
    const rifa = { rifaId: id, titulo: text_(p.titulo, 80), cantidad: quantity, precio: price,
      premio: text_(p.premio, 120), loteria: text_(p.loteria, 60), fecha: text_(p.fecha, 10),
      hora: text_(p.hora, 5), responsable: text_(p.responsable, 80), condiciones: text_(p.condiciones, 500) };
    if (!rifa.titulo || !rifa.premio || !rifa.loteria || !rifa.responsable || !/^\d{4}-\d{2}-\d{2}$/.test(rifa.fecha) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(rifa.hora)) fail_("Completa todos los detalles de la rifa.");
    const row = [id, META_NUMBER_ID, "meta", String(quantity), "", new Date().toISOString(),
      safeCell_(rifa.titulo), price, safeCell_(rifa.premio), safeCell_(rifa.loteria), rifa.fecha, rifa.hora,
      safeCell_(rifa.responsable), "", 0, "", safeCell_(rifa.condiciones)];
    table_("entries").appendRow(row); rows_("entries").push(row);
    CacheService.getScriptCache().remove("rifapro:catalog:v3");
    return { ok: true, rifaId: id, rifa: { ...rifa, revision: 0 } };
  });
}
function deleteRaffle_(id) {
  return locked_(() => {
    raffle_(id);
    for (const key of ["entries"]) {
      const rows = rows_(key), sheet = table_(key);
      // Bottom-up avoids shifted row indices.
      for (let i = rows.length - 1; i >= 0; i--) if (String(rows[i][0]) === String(id)) { sheet.deleteRow(i + 2); rows.splice(i, 1); }
    }
    CacheService.getScriptCache().remove(`public:${id}`);
    CacheService.getScriptCache().remove("rifapro:catalog:v3");
    return { ok: true };
  });
}
function updatePrize_(p) {
  return locked_(() => {
    const rifa = raffle_(p.rifaId), premio = text_(p.premio, 120), titulo = text_(p.titulo || rifa.titulo, 80);
    if (!premio || !titulo) fail_("Escribe el premio y el nombre de la rifa.");
    requestId_(p);
    const entries = rows_("entries");
    const metas = entries.map((row, i) => ({ row, index: i })).filter((item) => String(item.row[0]) === rifa.rifaId && String(item.row[1]) === META_NUMBER_ID);
    const current = metas[metas.length - 1];
    if (current && String(current.row[15]) === p.requestId) return { ok: true, rifaId: rifa.rifaId, rifa };
    if (Number(p.revision) !== rifa.revision) fail_("El premio cambió en otra sesión. Actualiza antes de editarlo.", "STALE");
    const revision = rifa.revision + 1;
    // Only prize/title fields change on existing records; customer, payment,
    // original price and number quantity stay intact.
    const affected = entries.map((row, i) => ({ row, i })).filter((item) => String(item.row[0]) === rifa.rifaId);
    affected.forEach(({ row, i }) => {
      row[6] = safeCell_(titulo); row[8] = safeCell_(premio);
      if (String(row[1]) === META_NUMBER_ID) { row[5] = new Date().toISOString(); row[14] = revision; row[15] = p.requestId; }
    });
    for (let i = 0; i < affected.length;) {
      const first = affected[i].i, values = [affected[i].row.slice(6, 9)]; i++;
      while (i < affected.length && affected[i].i === first + values.length) values.push(affected[i++].row.slice(6, 9));
      table_("entries").getRange(first + 2, 7, values.length, 3).setValues(values);
    }
    metas.forEach(({ row, index }) => table_("entries").getRange(index + 2, 1, 1, HEADERS.length).setValues([Array.from({ length: HEADERS.length }, (_, i) => row[i] ?? "")]));
    if (!current) {
      const row = [rifa.rifaId, META_NUMBER_ID, "meta", String(rifa.cantidad), "", new Date().toISOString(), safeCell_(titulo), rifa.precio, safeCell_(premio), safeCell_(rifa.loteria), rifa.fecha, rifa.hora, safeCell_(rifa.responsable), "", revision, p.requestId, safeCell_(rifa.condiciones)];
      table_("entries").appendRow(row); entries.push(row);
    }
    CacheService.getScriptCache().remove("rifapro:catalog:v3");
    return { ok: true, rifaId: rifa.rifaId, rifa: { ...rifa, premio, titulo, revision } };
  });
}

// Run once from the original spreadsheet's editor. Existing rows and original
// columns are preserved. The four new columns are appended, never inserted.
function prepararSistema() {
  locked_(() => {
    const db = db_();
    PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", db.getId());
    Object.keys(TABLES).forEach((key) => {
      const def = TABLES[key];
      let sheet = db.getSheetByName(def.name);
      if (!sheet) { sheet = db.insertSheet(def.name); sheet.appendRow(def.headers); }
      else {
        const current = sheet.getRange(1, 1, 1, def.headers.length).getValues()[0];
        // Stop rather than overwrite a custom/incompatible column.
        current.forEach((header, i) => { if (header && String(header) !== def.headers[i]) fail_(`La columna ${i + 1} de ${def.name} no coincide. No se modificaron tus datos.`); });
        const firstMissing = current.findIndex((header) => !header);
        if (firstMissing >= 0) sheet.getRange(1, firstMissing + 1, 1, def.headers.length - firstMissing).setValues([def.headers.slice(firstMissing)]);
      }
      sheet.setFrozenRows(1);
      if (key === "entries") sheet.getRange("A:B").setNumberFormat("@");
    });
  });
}
// Optional: run explicitly once to publish the supplied ticket. Edit the responsible
// person's spelling/date here before running; the reference does not include a year.
function crearRifaDeLaBoleta() {
  createRaffle_({ rifaId: "boyaca-31-octubre-2026", titulo: "Rifa de $500.000", premio: "$500.000",
    cantidad: 100, precio: 20000, loteria: "Boyacá", fecha: "2026-10-31", hora: "17:00",
    responsable: "Arelys B. Mendoza B.", condiciones: "Quien no haya pagado el número el día 31 a las 5:00 p. m. perderá el derecho al premio." });
}
