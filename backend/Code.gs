/* Google Apps Script V8. See README.md before deploying or migrating data. */
const PROTOCOL = 2;
const TABLES = {
  rifas: { name: "RifaPro_v2_Rifas", headers: ["rifaId", "json", "createdAt"] },
  numbers: { name: "RifaPro_v2_Numeros", headers: ["rifaId", "numeroId", "estado", "clienteNombre", "clienteTelefono", "paymentMethod", "updatedAt", "revision", "requestId"] },
  requests: { name: "RifaPro_v2_Solicitudes", headers: ["requestId", "action", "rifaId", "numeroId", "result"] }
};
const SESSION_SECONDS = 1800;

function doGet(event) {
  const p = event.parameter || {};
  // JSONP is read-only and contains public information only.
  let result;
  try {
    if (p.action === "getRifas") result = { ok: true, protocol: PROTOCOL, data: readRaffles_() };
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
      const rifa = raffle_(payload.rifaId);
      // Read once: summary and displayed numbers refer to the same snapshot.
      const records = numberRows_(rifa.rifaId).map((n) => n.record);
      return output_({ ok: true, data: records, summary: summary_(rifa, records) });
    }
    if (request.action === "createRaffle") return output_(createRaffle_(payload));
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
  return { ok: false, code: error.code || "INVALID", error: error.safe ? error.message : "No se pudo procesar la solicitud. Revisa el servicio de reservas." };
}
function fail_(message, code) { const error = new Error(message); error.safe = true; error.code = code || "INVALID"; throw error; }
function db_() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) fail_("El organizador debe configurar el servicio de reservas.");
  return SpreadsheetApp.openById(id);
}
function table_(key) {
  const sheet = db_().getSheetByName(TABLES[key].name);
  if (!sheet) fail_("El organizador debe ejecutar prepararSistema antes de publicar.");
  return sheet;
}
function rows_(key) {
  const sheet = table_(key);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, TABLES[key].headers.length).getValues();
}
function text_(value, max) { return String(value == null ? "" : value).trim().slice(0, max); }
function safeCell_(value) { const s = String(value == null ? "" : value); return /^[=+@\-]/.test(s) ? "'" + s : s; }
function unsheet_(value) { return String(value == null ? "" : value).replace(/^'(?=[=+@\-])/, ""); }
function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) fail_("Hay varias reservas en curso. Intenta de nuevo en unos segundos.", "BUSY");
  try { const result = fn(); SpreadsheetApp.flush(); return result; }
  finally { lock.releaseLock(); }
}
function readRaffles_() { return rows_("rifas").map((row) => JSON.parse(row[1])); }
function raffle_(id) {
  const r = readRaffles_().find((item) => item.rifaId === String(id));
  if (!r) fail_("Esta rifa ya no está disponible.", "NOT_FOUND");
  return r;
}
function numberRows_(id) {
  return rows_("numbers").map((row, i) => ({ row: i + 2, record: {
    rifaId: String(row[0]), numeroId: String(row[1]), estado: String(row[2]),
    clienteNombre: unsheet_(row[3]), clienteTelefono: unsheet_(row[4]), paymentMethod: String(row[5]),
    updatedAt: row[6] instanceof Date ? row[6].toISOString() : String(row[6]), revision: Number(row[7]) || 0, requestId: String(row[8])
  } })).filter((n) => n.record.rifaId === String(id));
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
  const values = [record.rifaId, record.numeroId, record.estado, safeCell_(record.clienteNombre), safeCell_(record.clienteTelefono),
    record.paymentMethod, record.updatedAt, record.revision, record.requestId];
  const sheet = table_("numbers");
  const target = sheet.getRange(row || sheet.getLastRow() + 1, 1, 1, values.length);
  target.setNumberFormat("@"); target.setValues([values]);
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
  table_("requests").appendRow([p.requestId, action, p.rifaId, String(p.numeroId), JSON.stringify(result)]);
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
  return locked_(() => {
    const rifa = raffle_(p.rifaId), number = validNumber_(p, rifa);
    const previous = oldRequest_(action, p); if (previous) return previous;
    const existing = numberRows_(rifa.rifaId).find((n) => Number(n.record.numeroId) === Number(number));
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
    return remember_(action, p, { ok: true });
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
    if (readRaffles_().some((r) => r.rifaId === id)) return { ok: true, rifaId: id };
    const quantity = Number(p.cantidad), price = Number(p.precio);
    if (!Number.isInteger(quantity) || quantity < 2 || quantity > 100000 || !Number.isSafeInteger(price) || price < 1 || price > 100000000) fail_("Revisa el precio y la cantidad de números.");
    const rifa = { rifaId: id, titulo: text_(p.titulo, 80), cantidad: quantity, precio: price,
      premio: text_(p.premio, 120), loteria: text_(p.loteria, 60), fecha: text_(p.fecha, 10),
      hora: text_(p.hora, 5), responsable: text_(p.responsable, 80), condiciones: text_(p.condiciones, 500) };
    if (!rifa.titulo || !rifa.premio || !rifa.loteria || !rifa.responsable || !/^\d{4}-\d{2}-\d{2}$/.test(rifa.fecha) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(rifa.hora)) fail_("Completa todos los detalles de la rifa.");
    table_("rifas").appendRow([id, JSON.stringify(rifa), new Date().toISOString()]);
    return { ok: true, rifaId: id };
  });
}
function deleteRaffle_(id) {
  return locked_(() => {
    raffle_(id);
    for (const key of ["numbers", "rifas"]) {
      const rows = rows_(key), sheet = table_(key);
      // Bottom-up avoids shifted row indices.
      for (let i = rows.length - 1; i >= 0; i--) if (String(rows[i][0]) === String(id)) sheet.deleteRow(i + 2);
    }
    CacheService.getScriptCache().remove(`public:${id}`);
    return { ok: true };
  });
}

// Run once in the editor after setting SPREADSHEET_ID and ADMIN_PIN in Script Properties.
// Only creates new v2 tabs; never overwrites legacy sheets or existing raffle data.
function prepararSistema() {
  locked_(() => {
    const db = db_();
    Object.keys(TABLES).forEach((key) => {
      const def = TABLES[key];
      if (!db.getSheetByName(def.name)) { const sheet = db.insertSheet(def.name); sheet.appendRow(def.headers); sheet.setFrozenRows(1); }
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
