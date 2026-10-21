/* Rifa Pro. Public data and admin data follow separate server contracts. */
"use strict";
const CONFIG = window.RIFA_CONFIG;
const developerPage = new URL(location.href).searchParams.get("panel") === "desarrollador";
const CACHE_KEY = "rifapro.public.v2";
const $ = (id) => document.getElementById(id);
const state = { rifas: [], activeId: null, records: [], summary: null, token: null,
  expiresAt: 0, protocol: false, online: false, refreshing: false, generation: 0,
  page: 1, selected: null, busy: false, installPrompt: null,
  capabilities: {}, pending: new Map(), pendingRaffles: new Map(), viewCache: new Map(),
  adminQueue: [], adminSending: false, writeEpoch: 0, prizeDraft: null,
  role: "visitor", ownerOrder: null, numberLabels: new Map(), entryLoading: false, globalPinLogin: false };
let toastTimer;
const money = (value) => new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(Number(value) || 0);
const escapeHTML = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
const activeRaffle = () => state.rifas.find((r) => r.id === state.activeId);
const hasSession = () => Boolean(state.token && Date.now() < state.expiresAt);
const isAdmin = () => hasSession() && state.role === "owner" && state.ownerOrder?.status === "active";
const isPlatform = () => hasSession() && state.role === "developer";
const uuid = () => crypto.randomUUID();

function normalizeRaffle(raw) {
  return { id: String(raw.id || raw.rifaId), titulo: String(raw.titulo || "Rifa"),
    precio: Number(raw.precio) || 0, cantidad: Math.min(100000, Math.max(2, Number(raw.cantidad) || 100)),
    premio: String(raw.premio || ""), loteria: String(raw.loteria || ""), fecha: String(raw.fecha || ""),
    hora: String(raw.hora || ""), responsable: String(raw.responsable || ""), condiciones: String(raw.condiciones || ""), revision: Number(raw.revision) || 0,
    independiente: Boolean(raw.independiente), expiresAt: raw.expiresAt || null, expired: Boolean(raw.expired) };
}
const raffleExpired = (r) => Boolean(r && (r.expired || (r.expiresAt && Date.now() >= Date.parse(r.expiresAt))));
function formatDate(raw, short = false) {
  const match = String(raw).match(/^(\d{4})-(\d{2})-(\d{2})/);
  const date = match ? new Date(+match[1], +match[2] - 1, +match[3]) : new Date(raw);
  if (Number.isNaN(date.getTime())) return raw || "Por confirmar";
  return new Intl.DateTimeFormat("es-CO", { day: "numeric", month: short ? "short" : "long", ...(short ? {} : {year:"numeric"}) }).format(date);
}
function formatTime(raw) {
  const match = String(raw).match(/^(\d{1,2}):(\d{2})/);
  if (!match) return raw || "Por confirmar";
  const h = Number(match[1]);
  return `${h % 12 || 12}:${match[2]} ${h >= 12 ? "p. m." : "a. m."}`;
}
function status(message, kind = "online") {
  $("connectionText").textContent = message;
  $("connection").dataset.state = kind;
  document.querySelector(".live-label").lastChild.textContent = kind === "online" ? " Disponibilidad en línea" : " Disponibilidad sin confirmar";
  document.querySelector(".live-label").title = message;
}
function toast(message) {
  clearTimeout(toastTimer); $("toast").textContent = message; $("toast").hidden = false;
  toastTimer = setTimeout(() => { $("toast").hidden = true; }, 4500);
}
function markup(id, html) {
  if (state.viewCache.get(id) === html) return;
  const element = $(id), focused = document.activeElement;
  const number = element.contains(focused) ? focused.dataset.number : null;
  element.innerHTML = html; state.viewCache.set(id, html);
  if (number) element.querySelector(`[data-number="${number}"]`)?.focus({ preventScroll: true });
}
const operationKey = (rifaId, number) => `${rifaId}:${Number(number)}`;
function syncOperations() {
  const operations = [...state.pending.values(), ...state.pendingRaffles.values()];
  const failed = operations.filter((op) => op.status === "failed").length;
  $("syncTray").hidden = !operations.length;
  $("syncTray").classList.toggle("has-error", Boolean(failed));
  $("syncMessage").textContent = failed ? `${failed} cambio${failed > 1 ? "s" : ""} por confirmar` : `Confirmando ${operations.length} cambio${operations.length > 1 ? "s" : ""}…`;
  $("retryChanges").hidden = !failed;
  $("adminSummary").setAttribute("aria-busy", String(operations.some((op) => op.admin)));
}
function mergeRaffles(remote) {
  const map = new Map(remote.map((r) => [r.id, r]));
  for (const [id, operation] of state.pendingRaffles) {
    if (operation.action === "deleteRifa") map.delete(id);
    else map.set(id, operation.optimistic);
  }
  return [...map.values()];
}
function applyRecord(rifaId, record) {
  if (!record || state.activeId !== rifaId) return;
  const normalized = { ...record, numeroId: String(record.numeroId).padStart(Math.max(2, String(activeRaffle()?.cantidad - 1).length), "0") };
  state.records = state.records.filter((r) => Number(r.numeroId) !== Number(record.numeroId)); state.records.push(normalized);
}
function notice(title, message, success = false) {
  $("noticeTitle").textContent = title; $("noticeMessage").textContent = message;
  $("noticeSymbol").textContent = success ? "✓" : "✧";
  if (!$("noticeDialog").open) $("noticeDialog").showModal();
}
function setTheme(dark) {
  document.body.classList.toggle("dark-mode", dark);
  $("themeToggle").setAttribute("aria-label", dark ? "Activar modo claro" : "Activar modo oscuro");
  try { localStorage.setItem("rifapro.theme", dark ? "dark" : "light"); } catch { /* Private browsing. */ }
}
function savePublicCache() {
  // Never persist names, phones, payment methods, summaries or admin tokens.
  try {
    const confirmed = new Map(state.rifas.map((r) => [r.id, r]));
    for (const [id, op] of state.pendingRaffles) { if (op.before) confirmed.set(id, op.before); else confirmed.delete(id); }
    localStorage.setItem(CACHE_KEY, JSON.stringify({ rifas: [...confirmed.values()] }));
  } catch { /* Storage may be full. */ }
}
function jsonp(action, params = {}) {
  return new Promise((resolve, reject) => {
    const callback = `rifa_${uuid().replaceAll("-", "")}`;
    const url = new URL(CONFIG.apiUrl);
    url.searchParams.set("action", action); url.searchParams.set("callback", callback);
    url.searchParams.set("t", Date.now());
    for (const [key, value] of Object.entries(params)) url.searchParams.set(key, value);
    const script = document.createElement("script");
    let timer;
    const cleanup = () => { clearTimeout(timer); delete window[callback]; script.remove(); };
    window[callback] = (data) => { cleanup(); resolve(data); };
    script.onerror = () => { cleanup(); reject(new Error("No pudimos conectar con el servidor.")); };
    timer = setTimeout(() => { cleanup(); reject(new Error("El servidor tardó demasiado. Intenta de nuevo.")); }, 16000);
    script.src = url.href; document.body.appendChild(script);
  });
}
async function api(action, payload = {}, privateRequest = false) {
  if (!navigator.onLine) throw new Error("Necesitas conexión para confirmar esta operación.");
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 25000);
  try {
    const response = await fetch(CONFIG.apiUrl, { method: "POST", redirect: "follow", cache: "no-store",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action, payload, ...(privateRequest ? { token: state.token } : {}) }), signal: controller.signal });
    if (!response.ok) throw new Error("El servidor no está disponible. Intenta de nuevo.");
    const result = await response.json();
    if (result.code === "UNAUTHORIZED") {
      logout(false); throw new Error("Tu sesión terminó. Ingresa de nuevo como administrador.");
    }
    if (!result.ok && result.code !== "CONFLICT") { const error = new Error(result.error || "No se pudo confirmar la operación."); error.code = result.code || "REJECTED"; throw error; }
    return result;
  } catch (error) {
    if (error.name === "AbortError" || error instanceof TypeError) throw new Error("No llegó la confirmación. Comprueba tu conexión y reintenta; conservamos tu solicitud para evitar duplicados.");
    throw error;
  } finally { clearTimeout(timer); }
}

async function refresh() {
  if (developerPage) { if (isPlatform()) await loadOrders(); return; }
  if (state.refreshing || document.hidden) return;
  if (!navigator.onLine) { state.online = false; state.entryLoading = false; status("Sin conexión · solo consulta", "offline"); renderCartones(); return; }
  state.refreshing = true;
  const generation = state.generation, epoch = state.writeEpoch;
  try {
    if (state.role === "owner" && hasSession()) {
      await refreshOwnerOrder();
      if (!isAdmin()) { render(); return; }
    }
    let snapshot, catalog;
    if (!isAdmin() && state.capabilities.snapshot === undefined) {
      const initial = await jsonp("getSnapshot", { rifaId: state.activeId || new URL(location.href).searchParams.get("rifa") || "" });
      if (initial.ok && Array.isArray(initial.rifas)) { snapshot = initial; catalog = initial.rifas; }
      else state.capabilities.snapshot = false;
    }
    if (!snapshot && state.capabilities.snapshot) {
      snapshot = isAdmin() ? await api("adminSnapshot", { rifaId: state.activeId }, true) : await jsonp("getSnapshot", { rifaId: state.activeId || new URL(location.href).searchParams.get("rifa") || "" });
      catalog = snapshot.rifas;
    } else if (!snapshot) {
      const result = await jsonp("getRifas", { rifaId: new URL(location.href).searchParams.get("rifa") || "" });
      if (!result.ok || !Array.isArray(result.data)) throw new Error("No pudimos consultar las rifas.");
      if (generation !== state.generation || epoch !== state.writeEpoch) return;
      state.protocol = result.protocol === 2; state.capabilities = result.capabilities || { snapshot: false };
      catalog = result.data;
      const requested = new URL(location.href).searchParams.get("rifa");
      const selected = catalog.find((r) => String(r.rifaId || r.id) === state.activeId) || catalog.find((r) => String(r.rifaId || r.id) === requested) || catalog[0];
      const selectedId = selected ? String(selected.rifaId || selected.id) : null;
      snapshot = selectedId ? (isAdmin() ? await api("adminSnapshot", { rifaId: selectedId }, true) : await jsonp("getRifaNumbers", { rifaId: selectedId })) : { ok: true, data: [], rifaId: null };
      if (!snapshot.rifaId) snapshot.rifaId = selectedId;
    }
    if (generation !== state.generation || epoch !== state.writeEpoch) return;
    if (!snapshot.ok || !Array.isArray(snapshot.data) || !Array.isArray(catalog)) throw new Error("No pudimos actualizar la disponibilidad.");
    state.protocol = snapshot.protocol === 2 || state.protocol;
    state.capabilities = snapshot.capabilities || state.capabilities;
    const requestedId = new URL(location.href).searchParams.get("rifa");
    if (!isAdmin()) catalog = catalog.filter((r) => requestedId && String(r.rifaId || r.id) === requestedId);
    state.rifas = mergeRaffles(catalog.map(normalizeRaffle));
    if (!state.rifas.some((r) => r.id === state.activeId)) { state.activeId = state.rifas.find((r) => r.id === snapshot.rifaId)?.id || state.rifas[0]?.id || null; state.page = 1; }
    const rifa = activeRaffle();
    if (rifa && rifa.id === snapshot.rifaId) {
      state.records = snapshot.data.map((n) => ({ numeroId: String(n.numeroId).padStart(Math.max(2, String(rifa.cantidad - 1).length), "0"),
        estado: isAdmin() ? n.estado : n.estado === "libre" ? "libre" : "apartado", clienteNombre: String(n.clienteNombre || ""), updatedAt: n.updatedAt,
        ...(isAdmin() ? { clienteTelefono: n.clienteTelefono, paymentMethod: n.paymentMethod, revision: n.revision } : {}) }));
      state.summary = isAdmin() ? snapshot.summary : null;
    } else if (!rifa) { state.records = []; state.summary = null; }
    state.online = true; status(state.protocol ? "Actualizado en línea" : "Servidor pendiente de actualizar", state.protocol ? "online" : "error");
    savePublicCache(); render();
  } catch (error) { if (generation === state.generation) { state.online = false; status("No se pudo actualizar", "error"); } }
  finally { state.refreshing = false; if (generation === state.generation) { state.entryLoading = false; renderCartones(); } if (generation !== state.generation) queueMicrotask(refresh); }
}
function render() { renderList(); renderTicket(); renderNumbers(); renderAdmin(); syncOperations(); if (typeof renderCartones === "function") renderCartones(); }
function renderList() {
  $("raffleCount").textContent = state.rifas.length;
  markup("listaRifas", state.rifas.length ? state.rifas.map((r) => `<div class="raffle-item ${r.id === state.activeId ? "active" : ""} ${state.pendingRaffles.has(r.id) ? "pending" : ""}"><button class="raffle-select" type="button" data-raffle="${escapeHTML(r.id)}" aria-pressed="${r.id === state.activeId}"><span class="raffle-symbol" aria-hidden="true">✧</span><span><strong>${escapeHTML(r.titulo)}</strong><small>${escapeHTML(formatDate(r.fecha, true))} · ${r.cantidad} números</small></span><span class="arrow" aria-hidden="true">↗</span></button>${isAdmin() ? `<button class="delete-button" type="button" data-delete="${escapeHTML(r.id)}" aria-label="Eliminar ${escapeHTML(r.titulo)}">×</button>` : ""}</div>`).join("") : '<p class="muted">Aún no hay rifas publicadas.</p>');
}
function renderTicket() {
  const r = activeRaffle();
  if (!r) {
    if (state.role === "owner" && state.ownerOrder && !isAdmin()) {
      markup("infoRifa", '<div class="empty-state"><span class="note-mark" aria-hidden="true">✧</span><h2>Tu cartón está en camino.</h2><p>Guarda el enlace de tu cartón, realiza el pago y envía el comprobante. Cuando se apruebe, podrás entrar desde Administrador con tu PIN de cuatro dígitos.</p></div>');
      return;
    }
    const requested = new URL(location.href).searchParams.get("rifa");
    markup("infoRifa", requested ? '<div class="empty-state"><span class="note-mark" aria-hidden="true">✧</span><h2>Este cartón no está disponible.</h2><p>Comprueba el enlace con la persona responsable. Los cartones nuevos estarán disponibles después de aprobar su pago.</p></div>' : '<div class="empty-state"><span class="note-mark" aria-hidden="true">✧</span><h2>Tu rifa. Tu propio espacio.</h2><p>Crea un cartón por $1.000 COP. Tendrás un enlace para tus participantes y un acceso privado para gestionar números y pagos. Si te invitaron a participar, abre el enlace de esa rifa.</p></div>');
    return;
  }
  const prize = r.premio || r.titulo;
  const digits = Math.max(2, String(r.cantidad - 1).length);
  const editor = isAdmin() ? '<button class="prize-edit" type="button" data-edit-prize><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true"><path d="m16 3 5 5-12 12-6 1 1-6zM13 6l5 5"/></svg>Editar premio</button>' + (state.pendingRaffles.has(r.id) ? '<span class="ticket-saving">Confirmando…</span>' : '') : '';
  markup("infoRifa", `<div class="ticket-main"><div><p class="eyebrow">${escapeHTML(r.titulo.toUpperCase())}</p><h2 class="ticket-title">RIFA DE</h2><p class="ticket-prize">${escapeHTML(prize)}</p><p class="ticket-subtitle">${r.cantidad === 100 ? 'Juega con las dos (2) últimas cifras' : `Participa con uno de los ${r.cantidad} números`} de la lotería.</p>${editor}</div><div class="ticket-seal" aria-hidden="true"><span>ELIGE TU</span><strong>${digits}</strong><small>${r.cantidad === 100 ? 'ÚLTIMAS CIFRAS' : 'CIFRAS'}</small></div></div><div class="ticket-divider"></div><div class="ticket-details"><div><span class="detail-label">LOTERÍA</span><span class="detail-value">${escapeHTML(r.loteria || "Por confirmar")}</span></div><div><span class="detail-label">FECHA DEL SORTEO</span><span class="detail-value">${escapeHTML(formatDate(r.fecha))}</span></div><div><span class="detail-label">RESPONSABLE</span><span class="detail-value">${escapeHTML(r.responsable || "Por confirmar")}</span></div></div><div class="ticket-bottom"><div><p><strong>Hora límite de pago: ${escapeHTML(formatTime(r.hora))}.</strong><br>${escapeHTML(r.condiciones || "Quien no haya pagado su número a la hora límite del día del sorteo perderá el derecho al premio.")}</p></div><div class="ticket-price"><small>VALOR POR NÚMERO</small><strong>${money(r.precio)}</strong></div></div>`);
}
function numberRecord(n) { return state.records.find((record) => Number(record.numeroId) === Number(n)); }
function renderNumbers() {
  const r = activeRaffle(), grid = $("dashboard");
  if (!r) { markup("dashboard", '<p class="empty-state muted">Selecciona una rifa para ver sus números.</p>'); grid.dataset.layout = ""; markup("paginacion", ""); return; }
  const query = $("numberSearch").value.trim(), digits = Math.max(2, String(r.cantidad - 1).length);
  const labelKey = r.cantidad + ":" + digits;
  if (!state.numberLabels.has(labelKey)) { state.numberLabels.clear(); state.numberLabels.set(labelKey, Array.from({ length: r.cantidad }, (_, i) => String(i).padStart(digits, "0"))); }
  const labels = state.numberLabels.get(labelKey), all = query ? labels.filter((n) => n.includes(query)) : labels;
  const size = Math.max(20, Math.min(200, CONFIG.pageSize || 100)), pages = Math.ceil(all.length / size);
  state.page = Math.min(Math.max(state.page, 1), Math.max(pages, 1));
  const visible = all.slice((state.page - 1) * size, state.page * size);
  const layout = [r.id, state.page, query, r.cantidad].join("|");
  if (grid.dataset.layout !== layout) {
    grid.replaceChildren(); state.viewCache.delete("dashboard");
    for (const n of visible) { const tile = document.createElement("button"); tile.type = "button"; tile.dataset.number = n; tile.innerHTML = '<span class="num"></span><span class="number-state"></span>'; tile.firstChild.textContent = n; grid.appendChild(tile); }
    grid.dataset.layout = layout;
    if (!visible.length) grid.innerHTML = '<p class="empty-state muted">No encontramos ese número en esta rifa.</p>';
  }
  const records = new Map(state.records.map((n) => [Number(n.numeroId), n]));
  for (const tile of grid.querySelectorAll("[data-number]")) {
    const n = tile.dataset.number, operation = state.pending.get(operationKey(r.id, n));
    const record = operation?.optimistic || records.get(Number(n));
    const occupied = record && record.estado !== "libre", paid = isAdmin() && record?.estado === "pagado";
    const method = paid && record.paymentMethod;
    const label = operation ? operation.status === "failed" ? "Por confirmar" : "Confirmando…" : occupied ? record.clienteNombre || (paid ? "Pagado" : "Reservado") : state.online && state.protocol ? "Elegir" : "Sin verificar";
    const className = ["numero", paid ? "pagado" : occupied ? "apartado" : "libre", operation ? "pending" : "", operation?.status === "failed" ? "failed" : ""].filter(Boolean).join(" ");
    if (tile.className !== className) tile.className = className;
    const text = tile.querySelector(".number-state"); if (text.textContent !== label) text.textContent = label;
    const aria = 'Número ' + n + ', ' + (operation ? label : occupied ? 'reservado por ' + (record.clienteNombre || 'un participante') : state.online ? 'libre' : 'sin verificar');
    if (tile.getAttribute("aria-label") !== aria) tile.setAttribute("aria-label", aria);
    tile.setAttribute("aria-busy", String(Boolean(operation && operation.status !== "failed")));
    let badge = tile.querySelector(".tile-payment");
    if (method && !operation) { if (!badge) { badge = document.createElement("span"); tile.appendChild(badge); } badge.className = 'tile-payment ' + (method === 'efectivo' ? 'cash' : 'transfer'); badge.textContent = method === 'efectivo' ? '▤' : '↗'; badge.title = method === 'efectivo' ? 'Efectivo' : 'Nequi / transferencia'; badge.setAttribute("aria-label", badge.title); }
    else badge?.remove();
  }
  markup("paginacion", pages > 1 ? '<button type="button" data-page="' + (state.page - 1) + '" ' + (state.page === 1 ? 'disabled' : '') + ' aria-label="Página anterior">‹</button><span>' + state.page + ' de ' + pages + '</span><button type="button" data-page="' + (state.page + 1) + '" ' + (state.page === pages ? 'disabled' : '') + ' aria-label="Página siguiente">›</button>' : "");
}
function renderAdmin() {
  const admin = isAdmin();
  document.querySelectorAll(".admin-only").forEach((el) => { el.hidden = !admin; });
  $("paidLegend").hidden = !admin;
  $("adminButtonText").textContent = admin ? "Cerrar sesión" : "Administrador";
  const summary = admin ? state.summary : null;
  $("adminSummary").hidden = !summary;
  for (const [id, label, value] of [["kpiLibres", "Libres", summary?.libres], ["kpiApartados", "Apartados", summary?.apartados], ["kpiPagados", "Pagados", summary?.pagados], ["kpiRecaudado", "Recaudado", summary ? money(summary.recaudado) : ""]]) {
    $(id).innerHTML = summary ? `<strong>${label}</strong>${escapeHTML(value)}` : "";
  }
  $("totalTransfer").textContent = summary ? money(summary.transferencia) : "";
  $("totalCash").textContent = summary ? money(summary.efectivo) : "";
  $("unclassifiedPayments").hidden = !summary?.sinClasificar;
  $("unclassifiedPayments").textContent = summary?.sinClasificar ? `${money(summary.sinClasificar)} en pagos anteriores sin medio registrado.` : "";
  const percent = summary?.meta ? Math.min(100, summary.recaudado / summary.meta * 100) : 0;
  $("barra").style.width = `${percent}%`;
  document.querySelector(".progress").setAttribute("aria-valuenow", String(Math.round(percent)));
  $("porcentajeTexto").textContent = summary ? `${percent.toFixed(1)}% de ${money(summary.meta)} recaudado` : "";
}
function switchRaffle(id) {
  if (id === state.activeId) return;
  state.generation++; state.activeId = id; state.page = 1; state.records = []; state.summary = null; state.online = false;
  $("numberSearch").value = "";
  const url = new URL(location.href); url.searchParams.set("rifa", id); history.replaceState(null, "", url);
  status("Actualizando disponibilidad", "loading"); render(); refresh();
}
function showConflict(number, record) {
  notice(`El ${number} ya tiene dueño`, `${record?.clienteNombre || "Otro participante"} ya agendó este número${record?.updatedAt ? ` el ${new Intl.DateTimeFormat("es-CO", {timeZone:"America/Bogota",day:"numeric",month:"short",hour:"2-digit",minute:"2-digit",second:"2-digit"}).format(new Date(record.updatedAt))}` : ""}. Elige otro de los números libres.`);
}
function openNumber(n) {
  const pending = state.pending.get(operationKey(state.activeId, n));
  if (pending) { if (pending.status === "failed") retryChanges(); else toast("Este número se está confirmando. Puedes elegir otro mientras tanto."); return; }
  if (state.pendingRaffles.get(state.activeId)?.action === "createRaffle") { toast("La rifa se está creando. Podrás registrar números en cuanto se confirme."); return; }
  const r = activeRaffle(); if (!r) return;
  if (raffleExpired(r)) { notice("Esta rifa ya venció", "El cartón cerró dos días después del sorteo. Sus registros quedan disponibles para consulta."); return; }
  if (!state.protocol) { notice("Estamos preparando las reservas", "El organizador debe actualizar el servicio de reservas. Mientras tanto puedes consultar los detalles de la rifa."); return; }
  if (!state.online || !navigator.onLine) { notice("Conéctate para reservar", "Necesitamos comprobar la disponibilidad en línea. Intenta de nuevo cuando se restablezca la conexión."); refresh(); return; }
  const record = numberRecord(n);
  if (record?.estado !== "libre" && record && !isAdmin()) { showConflict(n, record); return; }
  state.selected = { rifaId: r.id, numeroId: n, requestId: uuid(), revision: record?.revision || 0 };
  $("numberForm").reset(); $("selectedNumber").textContent = n; $("numberError").textContent = "";
  $("clienteNombre").value = isAdmin() ? record?.clienteNombre || "" : "";
  $("clienteTelefono").value = isAdmin() ? record?.clienteTelefono || "" : "";
  $("estadoSelect").value = record?.estado === "pagado" ? "pagado" : "apartado";
  if (record?.paymentMethod) document.querySelector(`input[name=paymentMethod][value="${record.paymentMethod === "efectivo" ? "efectivo" : "transferencia"}"]`).checked = true;
  $("stateField").hidden = !isAdmin(); $("releaseNumberButton").hidden = !isAdmin() || !record || record.estado === "libre";
  $("numberDialogTitle").textContent = isAdmin() ? "Gestionar este número" : "Un paso más para participar";
  $("numberHelp").textContent = isAdmin() ? "Actualiza los datos del participante y registra su pago." : "Completa tus datos y confirma tu reserva.";
  $("saveNumberButton").textContent = isAdmin() ? "Guardar cambios →" : "Confirmar mi reserva →";
  updatePaymentField(); $("numberDialog").showModal();
}
function updatePaymentField() {
  document.querySelectorAll(".segmented-control [data-state]").forEach((button) => button.setAttribute("aria-pressed", String(button.dataset.state === $("estadoSelect").value)));
  const paid = isAdmin() && $("estadoSelect").value === "pagado";
  $("paymentField").hidden = !paid;
  document.querySelectorAll("input[name=paymentMethod]").forEach((input) => { input.required = paid; });
}
function confirmAction(title, message, accept) {
  return new Promise((resolve) => {
    const dialog = $("confirmDialog"); let accepted = false;
    $("confirmTitle").textContent = title; $("confirmMessage").textContent = message; $("confirmAccept").textContent = accept;
    const onAccept = () => { accepted = true; dialog.close(); resolve(true); };
    $("confirmAccept").addEventListener("click", onAccept, { once: true });
    dialog.addEventListener("close", () => { $("confirmAccept").removeEventListener("click", onAccept); if (!accepted) resolve(false); }, { once: true });
    dialog.showModal();
  });
}
function localSummary() {
  if (!isAdmin() || !state.summary || !activeRaffle()) return;
  const r = activeRaffle(), paid = state.records.filter((n) => n.estado === "pagado"), reserved = state.records.filter((n) => n.estado === "apartado");
  const transfer = paid.filter((n) => n.paymentMethod === "transferencia").length * r.precio, cash = paid.filter((n) => n.paymentMethod === "efectivo").length * r.precio;
  state.summary = { libres: r.cantidad - paid.length - reserved.length, apartados: reserved.length, pagados: paid.length, recaudado: paid.length * r.precio, meta: r.cantidad * r.precio, transferencia: transfer, efectivo: cash, sinClasificar: paid.length * r.precio - transfer - cash };
}
function operationStore(op) { return op.kind === "number" ? state.pending : state.pendingRaffles; }
function operationCurrent(op) { return operationStore(op).get(op.key) === op && (!op.admin || (isAdmin() && op.session === state.token)); }
function settleOperation(op, result) {
  if (!operationCurrent(op)) return;
  state.writeEpoch++; operationStore(op).delete(op.key);
  if (result.code === "CONFLICT") { applyRecord(op.payload.rifaId, result.record); render(); showConflict(op.payload.numeroId, result.record); return; }
  if (op.kind === "number") {
    applyRecord(op.payload.rifaId, result.record || { ...op.optimistic, revision: Number(op.payload.revision || 0) + 1 });
    if (result.summary && state.activeId === op.payload.rifaId) state.summary = result.summary; else localSummary();
  } else if (op.action === "deleteRifa") {
    state.rifas = state.rifas.filter((r) => r.id !== op.payload.rifaId);
  } else {
    const rifa = normalizeRaffle(result.rifa || op.optimistic);
    state.rifas = state.rifas.filter((r) => r.id !== rifa.id); state.rifas.unshift(rifa);
    if (op.action === "createRaffle" && state.activeId === rifa.id) state.summary = { libres: rifa.cantidad, apartados: 0, pagados: 0, recaudado: 0, meta: rifa.cantidad * rifa.precio, transferencia: 0, efectivo: 0, sinClasificar: 0 };
  }
  render(); savePublicCache();
  if (!op.admin && !document.querySelector("dialog[open]")) notice('¡El ' + op.payload.numeroId + ' es tuyo!', 'Tu reserva quedó confirmada. Acuerda el pago con ' + (activeRaffle()?.responsable || 'la persona responsable') + '.', true);
  else toast(op.kind === "number" ? 'Número ' + op.payload.numeroId + (op.action === "releaseNumber" ? ' liberado.' : ' guardado.') : op.action === "createRaffle" ? 'Rifa creada. Ya puedes compartirla.' : op.action === "deleteRifa" ? 'Rifa eliminada.' : 'Premio actualizado.');
}
function failOperation(op, error) {
  if (!operationCurrent(op)) return;
  state.writeEpoch++;
  if (error.code && !["BUSY", "RETRY"].includes(error.code)) {
    operationStore(op).delete(op.key);
    if (op.kind === "raffle") { state.rifas = state.rifas.filter((r) => r.id !== op.payload.rifaId); if (op.before) state.rifas.unshift(op.before); if (!activeRaffle()) state.activeId = state.rifas[0]?.id || null; }
    toast(error.message); state.generation++; refresh();
  } else { op.status = "failed"; op.error = error.message; toast("No llegó la confirmación. Tu cambio está listo para reintentar."); }
  render();
}
async function sendOperation(op) {
  try { settleOperation(op, await api(op.action, op.payload, op.admin)); } catch (error) { failOperation(op, error); }
}
async function flushAdminQueue() {
  if (state.adminSending || !state.adminQueue.length || !isAdmin()) return;
  state.adminSending = true;
  const count = state.capabilities.batchNumbers ? 25 : 1, operations = state.adminQueue.splice(0, count);
  try {
    if (operations.length > 1) {
      const response = await api("batchNumbers", { changes: operations.map((op) => ({ action: op.action, payload: op.payload })) }, true);
      for (const op of operations) { const result = response.results?.find((r) => r.requestId === op.payload.requestId); if (!result) failOperation(op, new Error("No llegó la confirmación.")); else if (result.ok || result.code === "CONFLICT") settleOperation(op, result); else { const error = new Error(result.error || "No se pudo guardar."); error.code = result.code; failOperation(op, error); } }
    } else await sendOperation(operations[0]);
  } catch (error) { operations.forEach((op) => failOperation(op, error)); }
  finally { state.adminSending = false; if (state.adminQueue.length) queueMicrotask(flushAdminQueue); }
}
function submitOperation(op) {
  op.status = "pending"; op.session = state.token;
  operationStore(op).set(op.key, op); state.writeEpoch++; render();
  if (op.admin && op.kind === "number") { state.adminQueue.push(op); flushAdminQueue(); } else sendOperation(op);
}
function retryChanges() {
  for (const op of [...state.pending.values(), ...state.pendingRaffles.values()]) if (op.status === "failed") submitOperation(op);
}
function saveNumber(event) {
  event.preventDefault(); if (!state.selected) return;
  const nombre = $("clienteNombre").value.trim(), telefono = $("clienteTelefono").value.trim();
  const method = document.querySelector("input[name=paymentMethod]:checked")?.value || "", admin = isAdmin();
  if (nombre.length < 2 || telefono.replace(/\D/g, "").length < 7) { $("numberError").textContent = "Escribe un nombre y un celular válido (mínimo 7 dígitos)."; return; }
  if (admin && $("estadoSelect").value === "pagado" && !method) { $("numberError").textContent = "Elige Nequi / transferencia o efectivo."; return; }
  const selected = { ...state.selected }, key = operationKey(selected.rifaId, selected.numeroId);
  if (state.pending.has(key)) return;
  const estado = admin ? $("estadoSelect").value : "apartado";
  const payload = { ...selected, clienteNombre: nombre, clienteTelefono: telefono, ...(admin ? { estado, paymentMethod: estado === "pagado" ? method : "" } : {}) };
  state.selected = null; $("numberDialog").close();
  submitOperation({ kind: "number", key, action: admin ? "updateNumber" : "reserveNumber", payload, admin,
    optimistic: { numeroId: selected.numeroId, estado, clienteNombre: nombre, ...(admin ? { clienteTelefono: telefono, paymentMethod: payload.paymentMethod } : {}) } });
}
async function releaseNumber() {
  if (!isAdmin() || !state.selected) return;
  const selected = { ...state.selected };
  if (!await confirmAction('¿Liberar el ' + selected.numeroId + '?', 'El número volverá a estar disponible. Se eliminarán los datos del participante y el pago registrado.', 'Liberar número')) return;
  $("numberDialog").close(); state.selected = null;
  submitOperation({ kind: "number", key: operationKey(selected.rifaId, selected.numeroId), action: "releaseNumber", payload: selected, admin: true, optimistic: { numeroId: selected.numeroId, estado: "libre", clienteNombre: "" } });
}
async function login(event) {
  event.preventDefault();
  if (!state.protocol) { $("adminError").textContent = "El servicio debe actualizarse para habilitar el resumen privado."; return; }
  $("loginButton").disabled = true; $("loginButton").textContent = "Verificando…"; $("adminError").textContent = "";
  try {
    const rifaId = state.globalPinLogin ? undefined : state.activeId || new URL(location.href).searchParams.get("rifa") || undefined;
    const result = await api("login", { pin: $("pinInput").value, ...(rifaId ? { rifaId } : {}) });
    if (!result.token || !result.expiresAt) throw new Error("El servidor no confirmó la sesión.");
    if (result.role !== "owner" || !result.rifaId) throw new Error("El acceso de organizador debe actualizarse antes de continuar.");
    useOwnerSession(result);
    $("pinInput").value = ""; $("adminDialog").close(); state.entryLoading = true; render(); refresh();
    toast("Sesión de administrador iniciada.");
  } catch (error) { $("adminError").textContent = error.message; }
  finally { $("pinInput").value = ""; $("loginButton").disabled = false; $("loginButton").textContent = "Entrar al resumen →"; }
}
function logout(notify = true) {
  const token = state.token;
  if (typeof clearOwnerAccess === "function") clearOwnerAccess();
  for (const [id, operation] of state.pendingRaffles) { state.rifas = state.rifas.filter((r) => r.id !== id); if (operation.before) state.rifas.unshift(operation.before); }
  state.token = null; state.expiresAt = 0; state.role = "visitor"; state.ownerOrder = null; state.summary = null; state.selected = null; state.records = [];
  state.rifas = []; state.activeId = new URL(location.href).searchParams.get("rifa");
  state.entryLoading = !developerPage && Boolean(state.activeId);
  state.generation++; state.online = false; state.pending.clear(); state.pendingRaffles.clear(); state.adminQueue = []; state.prizeDraft = null;
  for (const id of ["numberDialog", "raffleDialog", "prizeDialog"]) $(id).close();
  $("numberForm").reset(); $("raffleForm").reset(); $("prizeForm").reset(); render(); refresh();
  if (token) api("logout", { sessionToken: token }).catch(() => {});
  if (notify) toast("Sesión cerrada. Estás en la vista de participante.");
}
function createRaffle(event) {
  event.preventDefault();
  if ($("raffleForm").dataset.mode === "paid") { createPaidCard(); return; }
  if (!isPlatform()) return;
  const payload = Object.fromEntries(["titulo", "precio", "cantidad", "premio", "loteria", "fecha", "hora", "responsable", "condiciones"].map((id) => [id, $(id).value.trim()]));
  $("raffleForm").dataset.requestId ||= uuid(); payload.rifaId = $("raffleForm").dataset.requestId;
  if (state.pendingRaffles.has(payload.rifaId)) return;
  const optimistic = normalizeRaffle(payload);
  state.rifas.unshift(optimistic); state.activeId = optimistic.id; state.page = 1; state.records = []; state.summary = null; state.generation++;
  $("raffleDialog").close();
  submitOperation({ kind: "raffle", key: payload.rifaId, action: "createRaffle", payload, optimistic, admin: true });
}
async function deleteRaffle(id) {
  if (!isAdmin() || state.pendingRaffles.has(id)) return;
  const before = state.rifas.find((r) => r.id === id); if (!before) return;
  if (!await confirmAction('¿Eliminar esta rifa?', 'Se eliminarán ' + before.titulo + ' y sus reservas. Esta acción no se puede deshacer.', 'Eliminar rifa')) return;
  state.rifas = state.rifas.filter((r) => r.id !== id);
  if (state.activeId === id) { state.activeId = state.rifas[0]?.id || null; state.records = []; state.summary = null; state.generation++; }
  submitOperation({ kind: "raffle", key: id, action: "deleteRifa", payload: { rifaId: id }, before, admin: true });
  refresh();
}
function openPrizeEditor() {
  const rifa = activeRaffle(); if (!isAdmin() || !rifa || state.pendingRaffles.has(rifa.id)) return;
  if (raffleExpired(rifa)) { toast("Este cartón venció y está disponible para consulta."); return; }
  if (!state.capabilities.editPrize) { notice("La edición está por habilitarse", "Aplica la versión actualizada del servicio para editar el premio de esta rifa."); return; }
  state.prizeDraft = { ...rifa };
  $("editPremio").value = rifa.premio; $("editTitulo").value = rifa.titulo; $("prizePreview").textContent = rifa.premio; $("prizeError").textContent = "";
  $("prizeDialog").showModal();
}
function savePrize(event) {
  event.preventDefault(); if (!isAdmin() || !state.prizeDraft) return;
  const before = state.prizeDraft, premio = $("editPremio").value.trim(), titulo = $("editTitulo").value.trim();
  if (!premio || !titulo) return;
  const optimistic = { ...before, premio, titulo }, payload = { rifaId: before.id, premio, titulo, revision: before.revision, requestId: uuid() };
  state.rifas = state.rifas.map((r) => r.id === before.id ? optimistic : r); $("prizeDialog").close();
  submitOperation({ kind: "raffle", key: before.id, action: "updatePrize", payload, optimistic, before, admin: true });
}
async function shareRaffle() {
  const r = activeRaffle(); if (!isAdmin() || !r) return;
  if (state.pendingRaffles.get(r.id)?.action === "createRaffle") { toast("Estamos confirmando la creación de tu rifa."); return; }
  const url = new URL(location.href); url.search = ""; url.hash = ""; url.searchParams.set("rifa", r.id);
  await copyText(url.href, "Enlace copiado. Envíalo a tus clientes; entrarán como participantes.");
}
function setupPWA() {
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("./sw.js").catch(() => {});
  window.addEventListener("beforeinstallprompt", (event) => { event.preventDefault(); state.installPrompt = event; $("installButton").hidden = false; });
  window.addEventListener("appinstalled", () => { state.installPrompt = null; $("installButton").hidden = true; toast("Rifa Pro ya está en tu inicio."); });
  const ios = /iPad|iPhone|iPod/.test(navigator.userAgent);
  if (ios && !window.matchMedia("(display-mode: standalone)").matches && !navigator.standalone) $("installButton").hidden = false;
  $("installButton").addEventListener("click", async () => {
    if (!state.installPrompt) { notice("Rifa Pro en tu inicio", "En Safari, toca Compartir y luego Agregar a inicio. El icono de tu rifa quedará siempre a mano."); return; }
    const prompt = state.installPrompt; state.installPrompt = null; await prompt.prompt(); await prompt.userChoice; $("installButton").hidden = true;
  });
}
function init() {
  try { const cached = JSON.parse(localStorage.getItem(CACHE_KEY)); if (Array.isArray(cached?.rifas)) state.rifas = cached.rifas.map(normalizeRaffle); setTheme(localStorage.getItem("rifapro.theme") === "dark"); } catch { /* Start clean. */ }
  const requested = new URL(location.href).searchParams.get("rifa");
  state.rifas = state.rifas.filter((r) => !developerPage && requested && r.id === requested);
  state.activeId = state.rifas.find((r) => r.id === requested)?.id || state.rifas[0]?.id || null;
  state.entryLoading = !developerPage && Boolean(requested);
  render(); status("Conectando", "loading");
  $("themeToggle").addEventListener("click", () => setTheme(!document.body.classList.contains("dark-mode")));
  $("adminButton").addEventListener("click", () => { if (hasSession()) logout(); else if (developerPage) $("developerPassword").focus(); else openPinLogin(!activeRaffle()); });
  $("adminForm").addEventListener("submit", login); $("numberForm").addEventListener("submit", saveNumber);
  $("raffleForm").addEventListener("submit", createRaffle); $("releaseNumberButton").addEventListener("click", releaseNumber);
  $("estadoSelect").addEventListener("change", updatePaymentField);
  $("newRaffleButton").addEventListener("click", () => { $("raffleForm").reset(); delete $("raffleForm").dataset.requestId; delete $("raffleForm").dataset.mode; $("creatorFields").hidden = true; $("creatorPhone").required = false; $("createRaffleButton").textContent = "Crear rifa →"; $("raffleError").textContent = ""; $("raffleDialog").showModal(); });
  $("numberSearch").addEventListener("input", () => { state.page = 1; renderNumbers(); });
  $("shareButton").addEventListener("click", shareRaffle);
  $("infoRifa").addEventListener("click", (event) => { if (event.target.closest("[data-edit-prize]")) openPrizeEditor(); });
  $("prizeForm").addEventListener("submit", savePrize);
  $("editPremio").addEventListener("input", () => { $("prizePreview").textContent = $("editPremio").value || "Tu próximo premio"; });
  $("retryChanges").addEventListener("click", retryChanges);
  document.querySelectorAll(".segmented-control [data-state]").forEach((button) => button.addEventListener("click", () => { $("estadoSelect").value = button.dataset.state; updatePaymentField(); }));
  $("listaRifas").addEventListener("click", (event) => { const button = event.target.closest("[data-raffle],[data-delete]"); if (button?.dataset.raffle) switchRaffle(button.dataset.raffle); else if (button?.dataset.delete) deleteRaffle(button.dataset.delete); });
  $("dashboard").addEventListener("click", (event) => { const n = event.target.closest("[data-number]")?.dataset.number; if (n !== undefined) openNumber(n); });
  $("paginacion").addEventListener("click", (event) => { const button = event.target.closest("[data-page]"); if (button && !button.disabled) { state.page = Number(button.dataset.page); renderNumbers(); $("numeros").scrollIntoView({ block: "start", behavior: "smooth" }); } });
  document.querySelectorAll("[data-close]").forEach((button) => button.addEventListener("click", () => { if (!state.busy) $(button.dataset.close).close(); }));
  $("numberDialog").addEventListener("cancel", (event) => { if (state.busy) event.preventDefault(); });
  $("numberDialog").addEventListener("close", () => { if (!$("numberDialog").open) state.selected = null; });
  window.addEventListener("offline", () => { state.online = false; state.entryLoading = false; status("Sin conexión · solo consulta", "offline"); renderNumbers(); renderCartones(); });
  window.addEventListener("online", () => { retryChanges(); refresh(); }); document.addEventListener("visibilitychange", () => { if (!document.hidden) { if (state.token && !hasSession()) logout(); refresh(); } });
  setupPWA(); refresh();
  setInterval(() => { if (state.token && !hasSession()) logout(); refresh(); }, Math.max(3000, CONFIG.pollInterval || 5000));
}
document.addEventListener("DOMContentLoaded", init);
