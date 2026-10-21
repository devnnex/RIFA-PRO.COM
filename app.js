/* Rifa Pro. Public data and admin data follow separate server contracts. */
"use strict";
const CONFIG = window.RIFA_CONFIG;
const CACHE_KEY = "rifapro.public.v2";
const $ = (id) => document.getElementById(id);
const state = { rifas: [], activeId: null, records: [], summary: null, token: null,
  expiresAt: 0, protocol: false, online: false, refreshing: false, generation: 0,
  page: 1, selected: null, busy: false, installPrompt: null };
let toastTimer;
const money = (value) => new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(Number(value) || 0);
const escapeHTML = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
const activeRaffle = () => state.rifas.find((r) => r.id === state.activeId);
const isAdmin = () => Boolean(state.token && Date.now() < state.expiresAt);
const uuid = () => crypto.randomUUID();

function normalizeRaffle(raw) {
  return { id: String(raw.id || raw.rifaId), titulo: String(raw.titulo || "Rifa"),
    precio: Number(raw.precio) || 0, cantidad: Math.min(100000, Math.max(2, Number(raw.cantidad) || 100)),
    premio: String(raw.premio || ""), loteria: String(raw.loteria || ""), fecha: String(raw.fecha || ""),
    hora: String(raw.hora || ""), responsable: String(raw.responsable || ""), condiciones: String(raw.condiciones || "") };
}
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
  try { localStorage.setItem(CACHE_KEY, JSON.stringify({ rifas: state.rifas })); } catch { /* Storage may be full. */ }
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
    if (!result.ok && result.code !== "CONFLICT") throw new Error(result.error || "No se pudo confirmar la operación.");
    return result;
  } catch (error) {
    if (error.name === "AbortError" || error instanceof TypeError) throw new Error("No llegó la confirmación. Comprueba tu conexión y reintenta; conservamos tu solicitud para evitar duplicados.");
    throw error;
  } finally { clearTimeout(timer); }
}

async function refresh() {
  if (state.refreshing || document.hidden) return;
  if (!navigator.onLine) { state.online = false; status("Sin conexión · solo consulta", "offline"); return; }
  state.refreshing = true;
  const generation = state.generation;
  try {
    const result = await jsonp("getRifas");
    if (!result.ok || !Array.isArray(result.data)) throw new Error("No pudimos consultar las rifas.");
    if (generation !== state.generation) return;
    state.protocol = result.protocol === 2;
    state.rifas = result.data.map(normalizeRaffle);
    if (!state.rifas.some((r) => r.id === state.activeId)) {
      const requested = new URL(location.href).searchParams.get("rifa");
      state.activeId = state.rifas.find((r) => r.id === requested)?.id || state.rifas[0]?.id || null;
      state.records = []; state.summary = null; state.page = 1;
    }
    const rifa = activeRaffle();
    if (rifa) {
      const snapshot = isAdmin() && state.protocol
        ? await api("adminSnapshot", { rifaId: rifa.id }, true)
        : await jsonp("getRifaNumbers", { rifaId: rifa.id });
      if (generation !== state.generation) return;
      if (!snapshot.ok || !Array.isArray(snapshot.data)) throw new Error("No pudimos actualizar la disponibilidad.");
      state.records = snapshot.data.map((n) => ({ numeroId: String(n.numeroId).padStart(Math.max(2, String(rifa.cantidad - 1).length), "0"),
        estado: isAdmin() ? n.estado : n.estado === "libre" ? "libre" : "apartado",
        clienteNombre: String(n.clienteNombre || ""), updatedAt: n.updatedAt,
        ...(isAdmin() ? { clienteTelefono: n.clienteTelefono, paymentMethod: n.paymentMethod, revision: n.revision } : {}) }));
      state.summary = isAdmin() ? snapshot.summary : null;
    }
    state.online = true;
    status(state.protocol ? "Actualizado en línea" : "Servidor pendiente de actualizar", state.protocol ? "online" : "error");
    savePublicCache(); render();
  } catch (error) {
    if (generation === state.generation) { state.online = false; status("No se pudo actualizar", "error"); }
  } finally {
    state.refreshing = false;
    // If a switch or logout occurred during the request, load the new view.
    if (generation !== state.generation) queueMicrotask(refresh);
  }
}
function render() { renderList(); renderTicket(); renderNumbers(); renderAdmin(); }
function renderList() {
  $("raffleCount").textContent = state.rifas.length;
  $("listaRifas").innerHTML = state.rifas.length ? state.rifas.map((r) => `<div class="raffle-item ${r.id === state.activeId ? "active" : ""}"><button class="raffle-select" type="button" data-raffle="${escapeHTML(r.id)}" aria-pressed="${r.id === state.activeId}"><span class="raffle-symbol" aria-hidden="true">✧</span><span><strong>${escapeHTML(r.titulo)}</strong><small>${escapeHTML(formatDate(r.fecha, true))} · ${r.cantidad} números</small></span><span class="arrow" aria-hidden="true">↗</span></button>${isAdmin() ? `<button class="delete-button" type="button" data-delete="${escapeHTML(r.id)}" aria-label="Eliminar ${escapeHTML(r.titulo)}">×</button>` : ""}</div>`).join("") : '<p class="muted">Aún no hay rifas publicadas.</p>';
}
function renderTicket() {
  const r = activeRaffle();
  if (!r) {
    $("infoRifa").innerHTML = '<div class="empty-state"><span class="note-mark" aria-hidden="true">✧</span><h2>La próxima ilusión está por llegar.</h2><p>Las rifas publicadas aparecerán aquí. El administrador puede crear una nueva desde su espacio.</p></div>';
    return;
  }
  const prize = r.premio || r.titulo;
  const digits = Math.max(2, String(r.cantidad - 1).length);
  $("infoRifa").innerHTML = `<div class="ticket-main"><div><p class="eyebrow">${escapeHTML(r.titulo.toUpperCase())}</p><h2 class="ticket-title">RIFA DE</h2><p class="ticket-prize">${escapeHTML(prize)}</p><p class="ticket-subtitle">${r.cantidad === 100 ? 'Juega con las dos (2) últimas cifras' : `Participa con uno de los ${r.cantidad} números`} de la lotería.</p></div><div class="ticket-seal" aria-hidden="true"><span>ELIGE TU</span><strong>${digits}</strong><small>${r.cantidad === 100 ? 'ÚLTIMAS CIFRAS' : 'CIFRAS'}</small></div></div><div class="ticket-divider"></div><div class="ticket-details"><div><span class="detail-label">LOTERÍA</span><span class="detail-value">${escapeHTML(r.loteria || "Por confirmar")}</span></div><div><span class="detail-label">FECHA DEL SORTEO</span><span class="detail-value">${escapeHTML(formatDate(r.fecha))}</span></div><div><span class="detail-label">RESPONSABLE</span><span class="detail-value">${escapeHTML(r.responsable || "Por confirmar")}</span></div></div><div class="ticket-bottom"><div><p><strong>Hora límite de pago: ${escapeHTML(formatTime(r.hora))}.</strong><br>${escapeHTML(r.condiciones || "Quien no haya pagado su número a la hora límite del día del sorteo perderá el derecho al premio.")}</p></div><div class="ticket-price"><small>VALOR POR NÚMERO</small><strong>${money(r.precio)}</strong></div></div>`;
}
function numberRecord(n) { return state.records.find((record) => Number(record.numeroId) === Number(n)); }
function renderNumbers() {
  const r = activeRaffle();
  $("paginacion").replaceChildren();
  if (!r) { $("dashboard").innerHTML = '<p class="empty-state muted">Selecciona una rifa para ver sus números.</p>'; return; }
  const query = $("numberSearch").value.trim();
  const digits = Math.max(2, String(r.cantidad - 1).length);
  const all = Array.from({ length: r.cantidad }, (_, i) => String(i).padStart(digits, "0")).filter((n) => !query || n.includes(query));
  const pageSize = Math.max(20, Math.min(200, CONFIG.pageSize || 100));
  const pages = Math.ceil(all.length / pageSize);
  state.page = Math.min(Math.max(state.page, 1), Math.max(pages, 1));
  const records = new Map(state.records.map((n) => [Number(n.numeroId), n]));
  $("dashboard").innerHTML = all.slice((state.page - 1) * pageSize, state.page * pageSize).map((n) => {
    const record = records.get(Number(n));
    const occupied = record && record.estado !== "libre";
    const paid = isAdmin() && record?.estado === "pagado";
    const payment = paid && record.paymentMethod;
    const label = occupied ? record.clienteNombre || (paid ? "Pagado" : "Reservado") : state.online && state.protocol ? "Elegir" : "Sin verificar";
    return `<button class="numero ${paid ? "pagado" : occupied ? "apartado" : "libre"}" type="button" data-number="${n}" aria-label="Número ${n}, ${occupied ? `reservado por ${escapeHTML(record.clienteNombre || "un participante")}` : "libre"}"><span class="num">${n}</span><span class="number-state">${escapeHTML(label)}</span>${payment ? `<span class="tile-payment ${payment === "efectivo" ? "cash" : "transfer"}" title="${payment === "efectivo" ? "Efectivo" : "Nequi / transferencia"}" aria-label="${payment === "efectivo" ? "Efectivo" : "Nequi / transferencia"}">${payment === "efectivo" ? "▤" : "↗"}</span>` : ""}</button>`;
  }).join("") || '<p class="empty-state muted">No encontramos ese número en esta rifa.</p>';
  if (pages > 1) $("paginacion").innerHTML = `<button type="button" data-page="${state.page - 1}" ${state.page === 1 ? "disabled" : ""} aria-label="Página anterior">‹</button><span>${state.page} de ${pages}</span><button type="button" data-page="${state.page + 1}" ${state.page === pages ? "disabled" : ""} aria-label="Página siguiente">›</button>`;
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
  if (state.busy) return;
  const r = activeRaffle(); if (!r) return;
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
  const paid = isAdmin() && $("estadoSelect").value === "pagado";
  $("paymentField").hidden = !paid;
  document.querySelectorAll("input[name=paymentMethod]").forEach((input) => { input.required = paid; });
}
async function saveNumber(event) {
  event.preventDefault();
  if (state.busy || !state.selected) return;
  const nombre = $("clienteNombre").value.trim();
  const telefono = $("clienteTelefono").value.trim();
  const method = document.querySelector("input[name=paymentMethod]:checked")?.value || "";
  if (nombre.length < 2 || telefono.replace(/\D/g, "").length < 7) { $("numberError").textContent = "Escribe un nombre y un celular válido (mínimo 7 dígitos)."; return; }
  const admin = isAdmin();
  if (admin && $("estadoSelect").value === "pagado" && !method) { $("numberError").textContent = "Elige Nequi / transferencia o efectivo para registrar el pago."; return; }
  const fingerprint = JSON.stringify([nombre, telefono, admin ? $("estadoSelect").value : "apartado", admin ? method : ""]);
  if (state.selected.fingerprint && state.selected.fingerprint !== fingerprint) state.selected.requestId = uuid();
  state.selected.fingerprint = fingerprint;
  const selection = { ...state.selected };
  const payload = { ...selection, clienteNombre: nombre, clienteTelefono: telefono,
    ...(admin ? { estado: $("estadoSelect").value, paymentMethod: $("estadoSelect").value === "pagado" ? method : "" } : {}) };
  state.busy = true; $("saveNumberButton").disabled = true; $("saveNumberButton").textContent = "Confirmando…";
  $("numberError").textContent = "";
  try {
    const result = await api(admin ? "updateNumber" : "reserveNumber", payload, admin);
    $("numberDialog").close();
    if (result.code === "CONFLICT") showConflict(selection.numeroId, result.record);
    else if (result.code === "STALE") notice("Este número cambió", "Otra sesión actualizó el número. Revisa los datos actuales antes de guardar.");
    else notice(admin ? "Cambios guardados" : `¡El ${selection.numeroId} es tuyo!`, admin ? "El registro quedó actualizado en el servidor." : `Tu reserva a nombre de ${result.record?.clienteNombre || nombre} quedó confirmada. Acuerda el pago con ${activeRaffle()?.responsable || "la persona responsable"}.`, true);
    state.selected = null; state.generation++; await refresh();
  } catch (error) { $("numberError").textContent = error.message; }
  finally { state.busy = false; $("saveNumberButton").disabled = false; $("saveNumberButton").textContent = isAdmin() ? "Guardar cambios →" : "Confirmar mi reserva →"; }
}
async function releaseNumber() {
  if (!isAdmin() || !state.selected || state.busy) return;
  if (!confirm(`¿Liberar el número ${state.selected.numeroId}? Se borrarán los datos y el pago registrado.`)) return;
  state.busy = true; $("releaseNumberButton").disabled = true;
  try {
    await api("releaseNumber", state.selected, true); $("numberDialog").close(); state.selected = null;
    state.generation++; await refresh(); toast("Número liberado.");
  } catch (error) { $("numberError").textContent = error.message; }
  finally { state.busy = false; $("releaseNumberButton").disabled = false; }
}
async function login(event) {
  event.preventDefault();
  if (!state.protocol) { $("adminError").textContent = "El servicio debe actualizarse para habilitar el resumen privado."; return; }
  $("loginButton").disabled = true; $("loginButton").textContent = "Verificando…"; $("adminError").textContent = "";
  try {
    const result = await api("login", { pin: $("pinInput").value });
    if (!result.token || !result.expiresAt) throw new Error("El servidor no confirmó la sesión.");
    state.token = result.token; state.expiresAt = result.expiresAt; state.generation++;
    $("pinInput").value = ""; $("adminDialog").close(); await refresh(); renderAdmin();
    toast("Sesión de administrador iniciada.");
  } catch (error) { $("adminError").textContent = error.message; }
  finally { $("pinInput").value = ""; $("loginButton").disabled = false; $("loginButton").textContent = "Entrar al resumen →"; }
}
function logout(notify = true) {
  const token = state.token;
  state.token = null; state.expiresAt = 0; state.summary = null; state.selected = null; state.records = [];
  state.generation++; state.online = false;
  for (const id of ["numberDialog", "raffleDialog"]) $(id).close();
  $("numberForm").reset(); $("raffleForm").reset(); render(); refresh();
  if (token) api("logout", { sessionToken: token }).catch(() => {});
  if (notify) toast("Sesión cerrada. Estás en la vista de participante.");
}
async function createRaffle(event) {
  event.preventDefault(); if (!isAdmin()) return;
  const payload = Object.fromEntries(["titulo", "precio", "cantidad", "premio", "loteria", "fecha", "hora", "responsable", "condiciones"].map((id) => [id, $(id).value.trim()]));
  // Keep the ID across network retries, so an uncertain response cannot duplicate a raffle.
  $("raffleForm").dataset.requestId ||= uuid(); payload.rifaId = $("raffleForm").dataset.requestId;
  $("createRaffleButton").disabled = true; $("raffleError").textContent = "";
  try { const result = await api("createRaffle", payload, true); $("raffleDialog").close(); state.activeId = result.rifaId; state.generation++; await refresh(); toast("Rifa creada. Ya puedes compartirla."); }
  catch (error) { $("raffleError").textContent = error.message; }
  finally { $("createRaffleButton").disabled = false; }
}
async function deleteRaffle(id) {
  if (!isAdmin() || !confirm("¿Eliminar esta rifa y todos sus números? Esta acción no se puede deshacer.")) return;
  try { await api("deleteRifa", { rifaId: id }, true); state.generation++; await refresh(); toast("Rifa eliminada."); }
  catch (error) { notice("No se pudo eliminar", error.message); }
}
async function shareRaffle() {
  const r = activeRaffle(); if (!r) { toast("Primero selecciona una rifa."); return; }
  const url = new URL(location.href); url.search = ""; url.hash = ""; url.searchParams.set("rifa", r.id);
  try {
    if (navigator.share) await navigator.share({ title: r.titulo, text: `Participa por ${r.premio}. Elige tu número.`, url: url.href });
    else { await navigator.clipboard.writeText(url.href); toast("Enlace de la rifa copiado."); }
  } catch (error) { if (error.name !== "AbortError") notice("Comparte esta rifa", url.href); }
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
  state.activeId = state.rifas.find((r) => r.id === requested)?.id || state.rifas[0]?.id || null;
  render(); status("Conectando", "loading");
  $("themeToggle").addEventListener("click", () => setTheme(!document.body.classList.contains("dark-mode")));
  $("adminButton").addEventListener("click", () => { if (isAdmin()) logout(); else { $("adminError").textContent = ""; $("adminDialog").showModal(); } });
  $("adminForm").addEventListener("submit", login); $("numberForm").addEventListener("submit", saveNumber);
  $("raffleForm").addEventListener("submit", createRaffle); $("releaseNumberButton").addEventListener("click", releaseNumber);
  $("estadoSelect").addEventListener("change", updatePaymentField);
  $("newRaffleButton").addEventListener("click", () => { $("raffleForm").reset(); delete $("raffleForm").dataset.requestId; $("raffleError").textContent = ""; $("raffleDialog").showModal(); });
  $("numberSearch").addEventListener("input", () => { state.page = 1; renderNumbers(); });
  $("shareButton").addEventListener("click", shareRaffle);
  $("listaRifas").addEventListener("click", (event) => { const button = event.target.closest("[data-raffle],[data-delete]"); if (button?.dataset.raffle) switchRaffle(button.dataset.raffle); else if (button?.dataset.delete) deleteRaffle(button.dataset.delete); });
  $("dashboard").addEventListener("click", (event) => { const n = event.target.closest("[data-number]")?.dataset.number; if (n !== undefined) openNumber(n); });
  $("paginacion").addEventListener("click", (event) => { const button = event.target.closest("[data-page]"); if (button && !button.disabled) { state.page = Number(button.dataset.page); renderNumbers(); $("numeros").scrollIntoView({ block: "start", behavior: "smooth" }); } });
  document.querySelectorAll("[data-close]").forEach((button) => button.addEventListener("click", () => { if (!state.busy) $(button.dataset.close).close(); }));
  $("numberDialog").addEventListener("cancel", (event) => { if (state.busy) event.preventDefault(); });
  $("numberDialog").addEventListener("close", () => { if (!state.busy) state.selected = null; });
  window.addEventListener("offline", () => { state.online = false; status("Sin conexión · solo consulta", "offline"); renderNumbers(); });
  window.addEventListener("online", refresh); document.addEventListener("visibilitychange", () => { if (!document.hidden) { if (state.token && !isAdmin()) logout(); refresh(); } });
  setupPWA(); refresh();
  setInterval(() => { if (state.token && !isAdmin()) logout(); refresh(); }, Math.max(3000, CONFIG.pollInterval || 5000));
}
document.addEventListener("DOMContentLoaded", init);
