/* Google Apps Script V8. Compatible with the original "rifas" spreadsheet. */
const PROTOCOL = 2;
const SHEET_NAME = "rifas";
const META_NUMBER_ID = "__META__";
const HEADERS = ["rifaId", "numeroId", "estado", "clienteNombre", "clienteTelefono", "updatedAt",
  "titulo", "precio", "premio", "loteria", "fecha", "hora", "responsable",
  "paymentMethod", "revision", "requestId", "condiciones"];
const TABLES = {
  entries: { name: SHEET_NAME, headers: HEADERS },
  requests: { name: "RifaPro_v2_Solicitudes", headers: ["requestId", "action", "rifaId", "numeroId", "result"] },
  orders: { name: "RifaPro_Cartones", headers: ["rifaId", "ownerHash", "estado", "createdAt", "updatedAt", "datos", "contacto", "importe"] },
  access: { name: "RifaPro_Accesos", headers: ["rifaId", "pinSalt", "pinHash", "failedAttempts", "lockedUntil", "pinLookup"] },
  winners: { name: "RifaPro_Ganadores", headers: ["rifaId", "numeroId", "nombre", "publishedAt", "revision", "requestId"] }
};
const SESSION_SECONDS = 1800;
const CAPABILITIES = { snapshot: true, editPrize: true, batchNumbers: true, writeSnapshot: true, independentRaffles: true, expiry: true, developerPanel: true, ownerPin: true, pinOnlyLogin: true, publicPayments: true, publishWinner: true };
const CARD_PRICE = 3500;
let principal = null;
let requestContext = null;
let stagedBatch = null;
function resetContext_() { requestContext = { db: null, sheets: {}, rows: {}, invalidate: new Set() }; }
function invalidateCache_(key) { requestContext.invalidate.add(key); }

function doGet(event) {
  resetContext_();
  const p = event.parameter || {};
  // JSONP is read-only and contains public information only.
  let result;
  try {
    if (p.action === "health") result = { ok: true, protocol: PROTOCOL, capabilities: CAPABILITIES, now: new Date().toISOString(), sheet: SHEET_NAME };
    else if (p.action === "getRifas") result = { ok: true, protocol: PROTOCOL, capabilities: CAPABILITIES, data: publicCatalog_(p.rifaId) };
    else if (p.action === "getSnapshot") {
      const rifas = publicCatalog_(p.rifaId);
      const rifa = rifas.find((r) => r.rifaId === p.rifaId) || rifas[0];
      result = { ok: true, protocol: PROTOCOL, capabilities: CAPABILITIES, rifas,
        rifaId: rifa ? rifa.rifaId : null, data: rifa ? publicSnapshot_(rifa.rifaId) : [] };
    }
    else if (p.action === "getRifaNumbers") {
      const rifa = publicRaffle_(p.rifaId);
      result = { ok: true, protocol: PROTOCOL, data: publicSnapshot_(rifa.rifaId) };
    } else result = { ok: false, error: "Acción pública no disponible. Actualiza la aplicación." };
  } catch (error) { result = failure_(error); }
  const callback = p.callback;
  if (callback && !/^rifa_[a-zA-Z0-9_]{1,80}$/.test(callback)) return output_({ ok: false, error: "Callback inválido." });
  return output_(result, callback);
}
function doPost(event) {
  resetContext_();
  principal = null;
  try {
    const request = JSON.parse(event.postData.contents);
    const payload = request.payload || {};
    if (request.action === "login") return output_(login_(payload.pin, payload.rifaId));
    if (request.action === "developerLogin") return output_(developerLogin_(payload.password));
    if (request.action === "orderCard") return output_(orderCard_(payload));
    if (request.action === "ownerLogin") return output_(ownerLogin_(payload));
    if (request.action === "logout") {
      const token = String(payload.sessionToken || "");
      if (/^[a-f0-9-]{60,100}$/.test(token)) CacheService.getScriptCache().remove(`session:${token}`);
      return output_({ ok: true });
    }
    if (request.action === "reserveNumber") return output_(reserve_(payload));
    principal = requireAdmin_(request.token);
    if (request.action === "developerDashboard") { platformOnly_(); return output_(developerDashboard_()); }
    if (request.action === "ownerStatus") return output_({ ok: true, order: ownerStatus_(principal) });
    if (request.action === "setOwnerPin") return output_(setOwnerPin_(payload));
    if (request.action === "publishWinner") return output_(publishWinner_(payload));
    if (request.action === "reportPayment") return output_(reportPayment_(principal));
    if (request.action === "listOrders") { platformOnly_(); return output_({ ok: true, orders: orders_().map(orderView_) }); }
    if (request.action === "approveCard") { platformOnly_(); return output_(approveCard_(payload)); }
    if (request.action === "adminSnapshot") {
      const rifas = readRaffles_().filter((r) => principal.role === "developer" || r.rifaId === principal.rifaId);
      if (principal.role === "owner") authorizeRaffle_(payload.rifaId || principal.rifaId, false);
      const rifa = rifas.find((r) => r.rifaId === payload.rifaId) || rifas[0];
      // Read once: summary and displayed numbers refer to the same snapshot.
      const records = rifa ? numberRows_(rifa.rifaId).map((n) => n.record) : [];
      return output_({ ok: true, protocol: PROTOCOL, capabilities: CAPABILITIES, rifas,
        rifaId: rifa ? rifa.rifaId : null, data: records, summary: rifa ? summary_(rifa, records) : null });
    }
    if (request.action === "createRaffle") { platformOnly_(); return output_(createRaffle_(payload)); }
    if (request.action === "updatePrize") return output_(updatePrize_(payload));
    if (request.action === "batchNumbers") return output_(batchNumbers_(payload));
    if (request.action === "updateNumber" || request.action === "releaseNumber") return output_(editNumber_(request.action, payload));
    if (request.action === "deleteRifa") { authorizeRaffle_(payload.rifaId, false); return output_(deleteRaffle_(payload.rifaId)); }
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
function prizeText_(value) {
  const text = text_(value, 120);
  const format = (amount) => "$" + (amount.replace(/\D/g, "").replace(/^0+(?=\d)/, "") || "0").replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return /^\$?\s*\d[\d.,\s]*$/.test(text) ? format(text) : text.replace(/\$\s*\d[\d.,]*/g, format);
}
function safeCell_(value) { const s = String(value == null ? "" : value); return /^[=+@\-]/.test(s) ? "'" + s : s; }
function unsheet_(value) { return String(value == null ? "" : value).replace(/^'(?=[=+@\-])/, ""); }
function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) fail_("Hay varias reservas en curso. Intenta de nuevo en unos segundos.", "BUSY");
  resetContext_();
  try {
    const result = fn(); SpreadsheetApp.flush();
    // Invalidate after Sheets flushes, so a concurrent public read cannot
    // repopulate a stale cache between the write and its confirmation.
    const cache = CacheService.getScriptCache(); requestContext.invalidate.forEach((key) => cache.remove(key));
    return result;
  }
  finally { lock.releaseLock(); }
}
function isoDate_(value) { return value instanceof Date ? value.toISOString() : String(value || ""); }
function legacyField_(value, format) {
  if (value instanceof Date) return Utilities.formatDate(value, "America/Bogota", format);
  return unsheet_(value);
}
function readRaffles_() {
  const groups = new Map();
  const winners = winnerRows_();
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
    r.expiresAt = expiry_(r.fecha);
    r.expired = r.expiresAt ? Date.now() >= Date.parse(r.expiresAt) : false;
    r.winner = winners.find((w) => w.rifaId === id)?.winner || null;
    return r;
  });
}
function publicCatalog_(id) {
  if (!id) return [];
  const cache = CacheService.getScriptCache(), key = "rifapro:catalog:v6";
  const cached = cache.get(key);
  let data = cached ? JSON.parse(cached) : null;
  if (!data) {
    const orders = orders_();
    data = readRaffles_().filter((r) => { const order = orders.find((o) => o.rifaId === r.rifaId); return !order || order.status === "active"; })
      .map((r) => ({ ...r, independiente: orders.some((o) => o.rifaId === r.rifaId) }));
    const json = JSON.stringify(data); if (json.length < 80000) cache.put(key, json, 5);
  }
  // All cards, including legacy cards, are accessible only by their own link.
  return data.filter((r) => r.rifaId === String(id)).map((r) => ({ ...r, expired: r.expiresAt ? Date.now() >= Date.parse(r.expiresAt) : false }));
}
function publicRaffle_(id) { const r = publicCatalog_(id)[0]; if (!r) fail_("Este cartón aún no está activo o ya no está disponible.", "NOT_FOUND"); return r; }
function winnerRows_() {
  // Existing raffles remain readable until prepararSistema creates the new table.
  if (!requestContext.sheets.winners) {
    if (requestContext.winnersMissing) return [];
    const sheet = db_().getSheetByName(TABLES.winners.name);
    if (!sheet) { requestContext.winnersMissing = true; return []; }
    requestContext.sheets.winners = sheet;
  }
  return rows_("winners").map((row, i) => ({ row: i + 2, rifaId: String(row[0]), requestId: String(row[5]),
    winner: { numeroId: String(row[1]), nombre: unsheet_(row[2]), publishedAt: isoDate_(row[3]), revision: Number(row[4]) || 0 } }));
}
function publishWinner_(p) {
  return locked_(() => {
    authorizeRaffle_(p.rifaId, false);
    const rifa = raffle_(p.rifaId), number = validNumber_(p, rifa), requestId = requestId_(p);
    const previous = winnerRows_().find((w) => w.rifaId === rifa.rifaId);
    if (previous?.requestId === requestId) return { ok: true, rifaId: rifa.rifaId, winner: previous.winner };
    if (Number(p.revision) !== (previous?.winner.revision || 0)) fail_("El ganador cambió en otra sesión. Actualiza antes de publicar.", "STALE");
    const record = numberRows_(rifa.rifaId).find((n) => Number(n.record.numeroId) === Number(number))?.record;
    if (!record || record.estado === "libre" || !record.clienteNombre.trim()) fail_("Elige un número asignado a un participante para publicar su nombre.");
    const winner = { numeroId: number, nombre: record.clienteNombre, publishedAt: new Date().toISOString(), revision: (previous?.winner.revision || 0) + 1 };
    const values = [rifa.rifaId, number, safeCell_(winner.nombre), winner.publishedAt, winner.revision, requestId];
    const sheet = table_("winners"), entries = rows_("winners"), row = previous?.row || entries.length + 2;
    const range = sheet.getRange(row, 1, 1, 6); range.setNumberFormat("@"); range.setValues([values]); entries[row - 2] = values;
    invalidateCache_("rifapro:catalog:v6");
    return { ok: true, rifaId: rifa.rifaId, winner };
  });
}
function expiry_(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) return null;
  // Available through the end of the second day after the draw, Colombia time.
  const time = Date.parse(date + "T00:00:00-05:00");
  return Number.isFinite(time) ? new Date(time + 3 * 86400000).toISOString() : null;
}
function assertOpen_(rifa) { if (rifa.expiresAt && Date.now() >= Date.parse(rifa.expiresAt)) fail_("Esta rifa venció dos días después del sorteo. Sus registros se conservan para consulta.", "EXPIRED"); }
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
  return { numeroId: record.numeroId, estado: ["libre", "apartado", "pagado"].includes(record.estado) ? record.estado : "apartado",
    paymentMethod: record.estado === "pagado" && ["transferencia", "efectivo"].includes(record.paymentMethod) ? record.paymentMethod : "",
    clienteNombre: record.estado === "libre" ? "" : record.clienteNombre, updatedAt: record.updatedAt };
}
function publicSnapshot_(id) {
  const cache = CacheService.getScriptCache();
  const key = `public:v3:${id}`;
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
  invalidateCache_(`public:v3:${record.rifaId}`);
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
    const rifa = publicRaffle_(p.rifaId), number = validNumber_(p, rifa), customer = validateCustomer_(p);
    assertOpen_(rifa);
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
    authorizeRaffle_(p.rifaId);
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
function verifyCredential_(value, expected, throttleKey) {
  const props = PropertiesService.getScriptProperties();
  let throttle = JSON.parse(props.getProperty(throttleKey) || "{}");
  if (throttle.until > Date.now()) fail_("Demasiados intentos. Espera unos minutos para volver a entrar.", "RATE_LIMIT");
  if (String(value || "") !== expected) {
    if (!throttle.start || Date.now() - throttle.start > 300000) throttle = { start: Date.now(), attempts: 0 };
    throttle.attempts = (throttle.attempts || 0) + 1;
    if (throttle.attempts >= 8) throttle.until = Date.now() + 300000;
    props.setProperty(throttleKey, JSON.stringify(throttle));
    fail_("Contraseña incorrecta.", "INVALID_PIN");
  }
  props.deleteProperty(throttleKey);
}
function newSession_(actor) {
  const token = Utilities.getUuid() + Utilities.getUuid(), expiresAt = Date.now() + SESSION_SECONDS * 1000;
  CacheService.getScriptCache().put("session:" + token, JSON.stringify({ ...actor, expiresAt }), SESSION_SECONDS);
  return { ok: true, token, expiresAt, ...actor };
}
function login_(pin, requestedId) {
  return locked_(() => {
    const props = PropertiesService.getScriptProperties(), scope = String(requestedId || "");
    const throttleKey = "OWNER_LOGIN_THROTTLE";
    const throttle = JSON.parse(props.getProperty(throttleKey) || "{}");
    if (!scope && throttle.until > Date.now()) fail_("Demasiados intentos. Espera cinco minutos para volver a entrar.", "RATE_LIMIT");
    const lookup = /^\d{4}$/.test(String(pin || "")) ? pinHash_(pin, "global-lookup-v1") : "";
    const matching = scope ? [accessRow_(scope)].filter(Boolean) : lookup ? rows_("access").map((row, index) => ({ row: index + 2, values: row }))
      .filter((a) => a.values[5] ? a.values[5] === lookup : pinHash_(pin, String(a.values[1])) === String(a.values[2])) : [];
    if (matching.length > 1) {
      matching.forEach((a) => { a.values[5] = lookup; table_("access").getRange(a.row, 1, 1, 6).setValues([a.values]); }); SpreadsheetApp.flush();
      fail_("Este PIN se repite en cartones anteriores. Contacta al desarrollador para actualizarlo.", "PIN_CONFLICT");
    }
    const access = matching[0];
    if (!access) {
      if (scope) fail_("Este cartón aún no tiene PIN configurado.", "PIN_NOT_SET");
      if (!throttle.start || Date.now() - throttle.start > 300000) { throttle.start = Date.now(); throttle.attempts = 0; }
      throttle.attempts = (throttle.attempts || 0) + 1; if (throttle.attempts >= 8) throttle.until = Date.now() + 300000;
      props.setProperty(throttleKey, JSON.stringify(throttle));
      fail_("PIN incorrecto. Revisa tus cuatro dígitos.", "INVALID_PIN");
    }
    const row = access.values, id = String(row[0]), until = Number(row[4]) || 0;
    if (until > Date.now()) fail_("Demasiados intentos. Espera cinco minutos para volver a entrar.", "RATE_LIMIT");
    if (until) { row[3] = 0; row[4] = 0; }
    if (!lookup || pinHash_(pin, String(row[1])) !== String(row[2])) {
      row[3] = (Number(row[3]) || 0) + 1; if (row[3] >= 8) row[4] = Date.now() + 300000;
      table_("access").getRange(access.row, 1, 1, row.length).setValues([row]); SpreadsheetApp.flush();
      fail_("PIN incorrecto para este cartón.", "INVALID_PIN");
    }
    row[3] = 0; row[4] = 0; row[5] = lookup;
    table_("access").getRange(access.row, 1, 1, 6).setValues([row]);
    const order = orders_().find((o) => o.rifaId === id);
    if (order && order.status !== "active") fail_("Tu cartón estará disponible cuando el desarrollador apruebe el pago.", "PAYMENT_PENDING");
    const rifa = raffle_(id), legacy = !order;
    return { ...newSession_({ role: "owner", rifaId: id, ...(legacy ? { legacy: true } : {}) }),
      capabilities: CAPABILITIES, order: order ? orderView_(order) : { rifaId: id, titulo: rifa.titulo, status: "active", legacy: true, hasPin: true } };
  });
}
function developerLogin_(password) {
  return locked_(() => {
    const props = PropertiesService.getScriptProperties(), expected = props.getProperty("DEVELOPER_PIN");
    if (!/^\d{4}$/.test(String(expected || ""))) fail_("Configura DEVELOPER_PIN con cuatro dígitos en las propiedades del servidor.");
    verifyCredential_(password, expected, "DEVELOPER_LOGIN_THROTTLE");
    return newSession_({ role: "developer" });
  });
}
function requireAdmin_(token) {
  if (!/^[a-f0-9-]{60,100}$/.test(String(token || ""))) fail_("Debes iniciar sesión.", "UNAUTHORIZED");
  const raw = CacheService.getScriptCache().get("session:" + token);
  // Previous global administrator sessions are revoked when roles are separated.
  const session = raw && raw[0] === "{" ? JSON.parse(raw) : null;
  if (!session || !["developer", "owner"].includes(session.role) || !session.expiresAt || session.expiresAt <= Date.now()) fail_("Tu sesión terminó. Ingresa de nuevo.", "UNAUTHORIZED");
  return session;
}
function createRaffle_(p) {
  return locked_(() => createRaffleLocked_(p));
}
function createRaffleLocked_(p) {
    const id = text_(p.rifaId, 80);
    if (!/^[a-zA-Z0-9_-]{1,80}$/.test(id)) fail_("Identificador de rifa inválido.");
    const existing = readRaffles_().find((r) => r.rifaId === id);
    if (existing) return { ok: true, rifaId: id, rifa: existing };
    const quantity = Number(p.cantidad), price = Number(p.precio);
    if (!Number.isInteger(quantity) || quantity < 2 || quantity > 100000 || !Number.isSafeInteger(price) || price < 1 || price > 100000000) fail_("Revisa el precio y la cantidad de números.");
    const rifa = { rifaId: id, titulo: text_(p.titulo, 80), cantidad: quantity, precio: price,
      premio: prizeText_(p.premio), loteria: text_(p.loteria, 60), fecha: text_(p.fecha, 10),
      hora: text_(p.hora, 5), responsable: text_(p.responsable, 80), condiciones: text_(p.condiciones, 500) };
    if (!rifa.titulo || !rifa.premio || !rifa.loteria || !rifa.responsable || !/^\d{4}-\d{2}-\d{2}$/.test(rifa.fecha) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(rifa.hora)) fail_("Completa todos los detalles de la rifa.");
    const row = [id, META_NUMBER_ID, "meta", String(quantity), "", new Date().toISOString(),
      safeCell_(rifa.titulo), price, safeCell_(rifa.premio), safeCell_(rifa.loteria), rifa.fecha, rifa.hora,
      safeCell_(rifa.responsable), "", 0, "", safeCell_(rifa.condiciones)];
    table_("entries").appendRow(row); rows_("entries").push(row);
    invalidateCache_("rifapro:catalog:v6");
    return { ok: true, rifaId: id, rifa: { ...rifa, revision: 0 } };
}
function deleteRaffle_(id) {
  return locked_(() => {
    if (!String(id || "").trim()) fail_("Selecciona la rifa que quieres eliminar.");
    // Retrying a deletion after a lost response is a success, not a restore.
    if (!readRaffles_().some((r) => r.rifaId === String(id))) return { ok: true };
    for (const key of ["entries", ...(db_().getSheetByName(TABLES.winners.name) ? ["winners"] : [])]) {
      const rows = rows_(key), sheet = table_(key);
      // Bottom-up avoids shifted row indices.
      for (let i = rows.length - 1; i >= 0; i--) if (String(rows[i][0]) === String(id)) { sheet.deleteRow(i + 2); rows.splice(i, 1); }
    }
    const order = orders_().find((o) => o.rifaId === String(id));
    if (order) { const row = rows_("orders")[order.row - 2]; row[2] = "deleted"; row[4] = new Date().toISOString(); table_("orders").getRange(order.row, 1, 1, 8).setValues([row]); }
    invalidateCache_(`public:v3:${id}`);
    invalidateCache_("rifapro:catalog:v6");
    return { ok: true };
  });
}
function updatePrize_(p) {
  return locked_(() => {
    authorizeRaffle_(p.rifaId);
    const rifa = raffle_(p.rifaId), premio = prizeText_(p.premio), titulo = text_(p.titulo || rifa.titulo, 80);
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
    invalidateCache_("rifapro:catalog:v6");
    return { ok: true, rifaId: rifa.rifaId, rifa: { ...rifa, premio, titulo, revision } };
  });
}

function platformOnly_() { if (principal?.role !== "developer") fail_("Esta acción corresponde al desarrollador de la plataforma.", "FORBIDDEN"); }
function developerDashboard_() {
  const orders = orders_(), rifas = readRaffles_(), accesses = rows_("access");
  const cards = orders.map((o) => {
    const current = rifas.find((r) => r.rifaId === o.rifaId);
    return { ...orderView_(o), titulo: current ? current.titulo : o.draft.titulo,
      expired: Date.now() >= Date.parse(expiry_(o.draft.fecha)), legacy: false };
  });
  rifas.filter((r) => !orders.some((o) => o.rifaId === r.rifaId)).forEach((r) => cards.push({ rifaId: r.rifaId, titulo: r.titulo, fecha: r.fecha, status: "active", expired: r.expired, legacy: true, amount: 0, contact: "", expiresAt: r.expiresAt }));
  cards.forEach((card) => { const access = accesses.find((a) => String(a[0]) === card.rifaId); card.pinConflict = Boolean(access?.[5] && accesses.filter((a) => a[5] === access[5]).length > 1); });
  return { ok: true, capabilities: CAPABILITIES, cards };
}
function orders_() {
  return rows_("orders").map((row, index) => ({ rifaId: String(row[0]), hash: String(row[1]), status: String(row[2]),
    createdAt: isoDate_(row[3]), updatedAt: isoDate_(row[4]), draft: JSON.parse(row[5]), contact: unsheet_(row[6]), amount: Number(row[7]), row: index + 2 }));
}
function orderView_(o) { return { rifaId: o.rifaId, reference: "RP-" + String(o.row - 1).padStart(5, "0"), status: o.status, createdAt: o.createdAt, titulo: o.draft.titulo, fecha: o.draft.fecha, contact: o.contact, amount: o.amount, expiresAt: expiry_(o.draft.fecha), hasPin: Boolean(accessRow_(o.rifaId)) }; }
function accessRow_(id) {
  const index = rows_("access").findIndex((row) => String(row[0]) === String(id));
  return index < 0 ? null : { row: index + 2, values: rows_("access")[index] };
}
function pinPepper_() {
  if (requestContext?.pinPepper) return requestContext.pinPepper;
  const props = PropertiesService.getScriptProperties();
  let pepper = props.getProperty("OWNER_PIN_PEPPER");
  if (!pepper) { pepper = Utilities.getUuid() + Utilities.getUuid(); props.setProperty("OWNER_PIN_PEPPER", pepper); }
  if (requestContext) requestContext.pinPepper = pepper;
  return pepper;
}
function pinHash_(pin, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ":" + String(pin) + ":" + pinPepper_(), Utilities.Charset.UTF_8)
    .map((b) => (b & 255).toString(16).padStart(2, "0")).join("");
}
function assertPinAvailable_(id, pin) {
  if (!/^\d{4}$/.test(String(pin || ""))) fail_("Elige un PIN de exactamente cuatro dígitos.", "INVALID_PIN_FORMAT");
  const lookup = pinHash_(pin, "global-lookup-v1");
  const occupied = rows_("access").some((row) => String(row[0]) !== String(id) && (row[5] ? row[5] === lookup : pinHash_(pin, String(row[1])) === String(row[2])));
  if (occupied || String(pin) === PropertiesService.getScriptProperties().getProperty("DEVELOPER_PIN")) fail_("Ese PIN ya está ocupado. Elige otros cuatro dígitos para tu cartón.", "PIN_IN_USE");
  return lookup;
}
function savePin_(id, pin) {
  if (accessRow_(id)) return;
  const lookup = assertPinAvailable_(id, pin), salt = Utilities.getUuid(), row = [id, salt, pinHash_(pin, salt), 0, 0, lookup];
  table_("access").appendRow(row); rows_("access").push(row);
}
function setOwnerPin_(p) {
  return locked_(() => {
    if (principal?.role !== "developer" && (principal?.role !== "owner" || String(p.rifaId) !== principal.rifaId)) fail_("Solo el dueño o el desarrollador pueden configurar este PIN.", "FORBIDDEN");
    const order = orders_().find((o) => o.rifaId === p.rifaId);
    if (order?.status === "deleted") fail_("Este cartón no está disponible.", "NOT_FOUND");
    const legacy = !order ? raffle_(p.rifaId) : null;
    const lookup = assertPinAvailable_(p.rifaId, p.pin), access = accessRow_(p.rifaId);
    if (access) {
      const salt = Utilities.getUuid(), row = [p.rifaId, salt, pinHash_(p.pin, salt), 0, 0, lookup];
      table_("access").getRange(access.row, 1, 1, 6).setValues([row]); requestContext.rows.access[access.row - 2] = row;
    } else savePin_(p.rifaId, p.pin);
    return { ok: true, order: order ? orderView_(order) : { rifaId: legacy.rifaId, titulo: legacy.titulo, status: "active", legacy: true, hasPin: true } };
  });
}
function secretHash_(secret) {
  if (!/^[a-f0-9-]{72}$/.test(String(secret || ""))) fail_("La clave privada del cartón no es válida.", "INVALID_KEY");
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret, Utilities.Charset.UTF_8).map((b) => (b & 255).toString(16).padStart(2, "0")).join("");
}
function ownerStatus_(actor) {
  if (actor.role === "owner" && actor.legacy) { const r = raffle_(actor.rifaId); return { rifaId: r.rifaId, titulo: r.titulo, status: "active", legacy: true }; }
  if (actor.role !== "owner") fail_("Ingresa con el enlace privado de tu cartón.", "FORBIDDEN");
  const order = orders_().find((o) => o.rifaId === actor.rifaId);
  if (!order) fail_("No encontramos este cartón.", "NOT_FOUND");
  return orderView_(order);
}
function authorizeRaffle_(id, writing = true) {
  if (!principal) fail_("Ingresa para gestionar tu rifa.", "UNAUTHORIZED");
  if (principal.role !== "developer") {
    if (String(id) !== principal.rifaId) fail_("Este cartón pertenece a otro organizador.", "FORBIDDEN");
    const order = orders_().find((o) => o.rifaId === String(id));
    if (principal.legacy ? Boolean(order) : (!order || order.status !== "active")) fail_("Tu cartón se activará después de revisar el comprobante de pago.", "PAYMENT_PENDING");
  }
  if (writing) assertOpen_(raffle_(id));
}
function orderCard_(p) {
  return locked_(() => {
    requestId_({ requestId: p.rifaId });
    const hash = secretHash_(p.ownerKey);
    const previous = orders_().find((o) => o.rifaId === p.rifaId);
    if (previous) { if (previous.hash !== hash) fail_("Este identificador ya pertenece a otro cartón.", "CONFLICT"); if (!accessRow_(p.rifaId)) savePin_(p.rifaId, p.adminPin); return { ok: true, order: orderView_(previous) }; }
    if (readRaffles_().some((r) => r.rifaId === p.rifaId)) fail_("Este identificador ya existe.", "CONFLICT");
    const contact = text_(p.contact, 30);
    if (contact.replace(/\D/g, "").length < 10) fail_("Escribe un celular válido para identificar tu comprobante.");
    const draft = { rifaId: p.rifaId };
    for (const [key, max] of [["titulo",80],["premio",120],["loteria",60],["fecha",10],["hora",5],["responsable",80],["condiciones",500]]) draft[key] = text_(p[key], max);
    draft.cantidad = Number(p.cantidad); draft.precio = Number(p.precio);
    draft.premio = prizeText_(draft.premio);
    if (!Number.isInteger(draft.cantidad) || draft.cantidad < 2 || draft.cantidad > 100000 || !Number.isSafeInteger(draft.precio) || draft.precio < 1 || draft.precio > 100000000) fail_("Revisa el precio y la cantidad de números.");
    if (!draft.titulo || !draft.premio || !draft.loteria || !draft.responsable || !/^\d{4}-\d{2}-\d{2}$/.test(draft.fecha) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(draft.hora)) fail_("Completa los detalles de la rifa.");
    const draw = Date.parse(draft.fecha + "T00:00:00-05:00");
    if (!Number.isFinite(draw) || new Date(draw).toISOString().slice(0, 10) !== draft.fecha) fail_("La fecha del sorteo no es válida.");
    assertOpen_({ expiresAt: expiry_(draft.fecha) });
    assertPinAvailable_(p.rifaId, p.adminPin);
    const now = new Date().toISOString(), row = [p.rifaId, hash, "pending", now, now, JSON.stringify(draft), safeCell_(contact), CARD_PRICE];
    table_("orders").appendRow(row); rows_("orders").push(row);
    savePin_(p.rifaId, p.adminPin);
    return { ok: true, order: orderView_(orders_().find((o) => o.rifaId === p.rifaId)) };
  });
}
function ownerLogin_(p) {
  const hash = secretHash_(p.ownerKey), order = orders_().find((o) => o.rifaId === String(p.rifaId) && o.hash === hash);
  if (!order) fail_("El enlace o la clave privada no corresponden a este cartón.", "INVALID_KEY");
  const token = Utilities.getUuid() + Utilities.getUuid(), expiresAt = Date.now() + SESSION_SECONDS * 1000;
  CacheService.getScriptCache().put(`session:${token}`, JSON.stringify({ role: "owner", rifaId: order.rifaId, expiresAt }), SESSION_SECONDS);
  return { ok: true, token, expiresAt, role: "owner", rifaId: order.rifaId, order: orderView_(order) };
}
function reportPayment_(actor) {
  if (actor.role !== "owner") fail_("Ingresa con el enlace privado de tu cartón.", "FORBIDDEN");
  return locked_(() => {
    const order = orders_().find((o) => o.rifaId === actor.rifaId);
    if (!order) fail_("No encontramos este cartón.");
    if (order.status === "pending") { const row = rows_("orders")[order.row - 2]; row[2] = "reviewing"; row[4] = new Date().toISOString(); table_("orders").getRange(order.row, 1, 1, 8).setValues([row]); }
    return { ok: true, order: orderView_(orders_().find((o) => o.rifaId === actor.rifaId)) };
  });
}
function approveCard_(p) {
  return locked_(() => {
    const order = orders_().find((o) => o.rifaId === String(p.rifaId));
    if (!order) fail_("No encontramos la solicitud.", "NOT_FOUND");
    if (order.status === "deleted") fail_("Este cartón fue eliminado y no se puede reactivar.", "NOT_FOUND");
    assertOpen_({ expiresAt: expiry_(order.draft.fecha) });
    if (order.status === "active") return { ok: true, order: orderView_(order) };
    // Orders keep the price quoted by this server when they were created.
    if (![1000, CARD_PRICE].includes(order.amount)) fail_("El valor del cartón no coincide.");
    createRaffleLocked_(order.draft);
    const row = rows_("orders")[order.row - 2]; row[2] = "active"; row[4] = new Date().toISOString();
    table_("orders").getRange(order.row, 1, 1, 8).setValues([row]);
    invalidateCache_("rifapro:catalog:v6");
    return { ok: true, order: orderView_(orders_().find((o) => o.rifaId === order.rifaId)) };
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
    pinPepper_();
    // Import the previous single-card PIN once. All new PINs live in the access table.
    const props = PropertiesService.getScriptProperties(), legacyPin = props.getProperty("ADMIN_PIN");
    if (/^\d{4}$/.test(String(legacyPin || ""))) {
      const orders = orders_(), legacy = readRaffles_().filter((r) => !orders.some((o) => o.rifaId === r.rifaId));
      const id = props.getProperty("LEGACY_RAFFLE_ID"), rifa = id ? legacy.find((r) => r.rifaId === id) : legacy.length === 1 ? legacy[0] : null;
      if (rifa) savePin_(rifa.rifaId, legacyPin);
    }
  });
}
// Optional: run explicitly once to publish the supplied ticket. Edit the responsible
// person's spelling/date here before running; the reference does not include a year.
function crearRifaDeLaBoleta() {
  createRaffle_({ rifaId: "boyaca-31-octubre-2026", titulo: "Rifa de $500.000", premio: "$500.000",
    cantidad: 100, precio: 20000, loteria: "Boyacá", fecha: "2026-10-31", hora: "17:00",
    responsable: "Arelys B. Mendoza B.", condiciones: "Quien no haya pagado el número el día 31 a las 5:00 p. m. perderá el derecho al premio." });
  prepararSistema(); // Import the previous PIN if this was the first legacy raffle.
}
