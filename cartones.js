/* Independent cards: no payment is considered verified by a client-side click. */
"use strict";
const PAYMENT_KEY = "3208759798";
const RECEIPT_PHONE = "573246394689";
let cardDraft = null;
let privateCardLink = "";
let ordersLoading = false;
let pinTarget = null;
const paymentLabels = { pending: "Pendiente de pago", reviewing: "Comprobante por revisar", active: "Cartón activo", deleted: "Cartón eliminado" };
function cardReference(order) { return order.reference || "RP-" + String(order.rifaId).replace(/-/g, "").slice(-6).toUpperCase(); }
function openPinLogin(global = true) {
  state.globalPinLogin = global; $("adminError").textContent = ""; $("pinInput").value = "";
  $("adminTitle").textContent = global ? "Vuelve a tu cartón." : "Tu rifa, en detalle.";
  $("adminDialog").showModal();
}

function scrollToNumbers() {
  $("numeros").scrollIntoView({ block: "start", behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "instant" : "smooth" });
  $("numbersTitle").tabIndex = -1; $("numbersTitle").focus({ preventScroll: true });
}
function renderCartones() {
  $("newRaffleButton").hidden = true;
  $("adminButtonText").textContent = hasSession() ? "Cerrar sesión" : "Administrador";
  $("adminButton").classList.toggle("organizer-access", !developerPage);
  $("publicCreateButton").hidden = hasSession();
  document.querySelector(".creator-entry").hidden = hasSession();
  $("ownerCard").hidden = !(hasSession() && state.role === "owner" && !state.ownerOrder?.legacy);
  if (state.ownerOrder) {
    $("ownerCardTitle").textContent = state.ownerOrder.titulo;
    $("ownerCardStatus").textContent = paymentLabels[state.ownerOrder.status] || "Por confirmar";
    $("cardPaymentStatus").textContent = paymentLabels[state.ownerOrder.status] || "Por confirmar";
    $("paidWhatsApp").hidden = ["active", "deleted"].includes(state.ownerOrder.status);
    $("enterMyCard").hidden = state.ownerOrder.status !== "active";
  }
  $("developerPanel").hidden = !developerPage;
  $("developerLoginForm").hidden = isPlatform();
  $("ordersPanel").hidden = !(developerPage && isPlatform());
  if ($("ordersPanel").hidden) {
    for (const id of ["ordersList", "developerStats"]) { if ($(id).childNodes.length) $(id).replaceChildren(); state.viewCache.delete(id); }
  }
  document.querySelector(".workspace").hidden = developerPage;
  document.querySelector(".page-intro").hidden = developerPage;
  if (developerPage) status(isPlatform() ? "Panel del desarrollador" : "Acceso privado", "online");
  $("entryLoader").hidden = !state.entryLoading || developerPage;
  $("numeros").hidden = !activeRaffle() || state.entryLoading;
  $("infoRifa").hidden = state.entryLoading;
  $("pinSetupButton").hidden = !isAdmin() || !state.capabilities.pinOnlyLogin;
  $("pinSetupButton").textContent = state.ownerOrder?.hasPin === false ? "Crear mi PIN" : "Cambiar mi PIN";
  document.body.classList.toggle("landing-page", !developerPage && !activeRaffle() && !hasSession());
  document.querySelector(".sidebar-heading").hidden = !activeRaffle();
  $("listaRifas").hidden = !activeRaffle();
  $("shareButton").hidden = !isAdmin() || !activeRaffle();
  const r = activeRaffle();
  $("expiryNotice").hidden = !r?.expiresAt;
  if (r?.expiresAt) $("expiryNotice").textContent = raffleExpired(r) ? "Cartón vencido · Los registros se conservan para consulta." : "Disponible hasta el " + new Intl.DateTimeFormat("es-CO", { day: "numeric", month: "long", year: "numeric", timeZone: "America/Bogota" }).format(new Date(Date.parse(r.expiresAt) - 1)) + " · Dos días después del sorteo.";
  $("chooseNumbers").hidden = !r;
  $("chooseNumbers").disabled = raffleExpired(r);
  if (hasSession() && state.role === "owner" && !isAdmin()) status("Cartón pendiente de activación", "online");
  if (typeof renderExperience === "function") renderExperience();
}
function startPaidCard() {
  if (hasSession()) { toast("Cierra tu sesión antes de crear otro cartón independiente."); return; }
  if (!state.capabilities.pinOnlyLogin) { notice("Estamos habilitando los nuevos cartones", "La creación de nuevos cartones estará disponible cuando el servicio esté listo. Intenta de nuevo más tarde."); return; }
  $("raffleForm").reset(); $("raffleForm").dataset.mode = "paid";
  $("raffleDialogTitle").textContent = "Tu rifa, tu propio espacio.";
  $("creatorFields").hidden = false; $("creatorPhone").required = true;
  $("creatorPin").required = true; $("creatorPinConfirm").required = true;
  $("createRaffleButton").textContent = "Continuar al pago · $3.500 →";
  $("raffleError").textContent = "";
  resetCreation();
  if (!cardDraft) cardDraft = { rifaId: uuid(), ownerKey: uuid() + uuid() };
  $("raffleDialog").showModal();
}
function rafflePublicLink(id) { const url = new URL(location.href); url.search = ""; url.hash = ""; url.searchParams.set("rifa", id); return url; }
function cardOwnerLink(id, key) { const url = rafflePublicLink(id); url.hash = "acceso=" + key; return url.href; }
function currentCardLink() { return state.ownerOrder?.hasPin ? rafflePublicLink(state.ownerOrder.rifaId).href : privateCardLink; }
function fillPayment(order) {
  const fee = currencyText(order.amount || 3500);
  $("cardFeeAmount").textContent = fee; $("paymentTransferTitle").textContent = "Transfiere " + fee;
  $("qrPaymentAmount").textContent = "Elkin Rangel · " + fee + " COP";
  $("paymentReference").textContent = cardReference(order);
  $("cardPaymentStatus").textContent = paymentLabels[order.status] || "Por confirmar";
  $("privateAccessLink").value = order.hasPin ? rafflePublicLink(order.rifaId).href : privateCardLink;
  $("privateAccessSection").hidden = !$("privateAccessLink").value;
  $("paidWhatsApp").href = "https://wa.me/" + RECEIPT_PHONE + "?text=" + encodeURIComponent("Hola, te escribo porque ya pagué mi cartón de rifa virtual. Te dejo el comprobante aquí:\n\nCartón: " + order.titulo + "\nReferencia: " + cardReference(order) + "\nValor: " + fee + " COP.\nAdjuntaré el comprobante en este chat.");
  $("paidWhatsApp").hidden = order.status === "active" || order.status === "deleted";
  $("enterMyCard").hidden = order.status !== "active";
  if (!$("paymentDialog").open) $("paymentDialog").showModal();
}
async function createPaidCard() {
  const button = $("createRaffleButton"); if (button.disabled) return;
  if (!validateCreation()) return;
  const pin = $("creatorPin").value;
  if (!/^\d{4}$/.test(pin) || pin !== $("creatorPinConfirm").value) { $("raffleError").textContent = "Escribe el mismo PIN de cuatro dígitos en ambos campos."; return; }
  formatMoneyField($("premio"));
  const payload = { ...cardDraft, adminPin: pin, contact: $("creatorPhone").value.trim(), ...Object.fromEntries(["titulo", "precio", "cantidad", "premio", "loteria", "fecha", "hora", "responsable", "condiciones"].map((id) => [id, $(id).value.trim()])) };
  payload.precio = amountValue(payload.precio);
  // Keep this recovery link before sending: an uncertain result can be retried safely.
  privateCardLink = cardOwnerLink(payload.rifaId, payload.ownerKey);
  try { sessionStorage.setItem("rifapro.cardAccess", privateCardLink); } catch { /* User can download the link. */ }
  button.disabled = true; button.textContent = "Preparando tu cartón…"; $("raffleError").textContent = "";
  try {
    const result = await api("orderCard", payload);
    const session = await api("ownerLogin", { rifaId: payload.rifaId, ownerKey: payload.ownerKey });
    useOwnerSession(session); $("raffleDialog").close(); $("creatorPin").value = ""; $("creatorPinConfirm").value = ""; fillPayment(result.order); cardDraft = null;
  } catch (error) { $("raffleError").textContent = error.message + " Puedes reintentar sin crear otro cartón."; }
  finally { button.disabled = false; button.textContent = "Continuar al pago · $3.500 →"; }
}
function useOwnerSession(result) {
  if (result.role !== "owner" || !result.rifaId) throw new Error("El servidor no confirmó el acceso a este cartón.");
  state.token = result.token; state.expiresAt = result.expiresAt; state.role = "owner"; state.ownerOrder = result.order;
  state.activeId = result.rifaId; state.rifas = []; state.records = []; state.summary = null;
  history.replaceState(null, "", rafflePublicLink(result.rifaId).href);
  state.capabilities = result.capabilities || state.capabilities;
  state.generation++; state.online = true; state.capabilities.independentRaffles = true; state.capabilities.snapshot = true;
  render();
}
async function refreshOwnerOrder() {
  if (state.ownerOrder?.status === "active") return;
  const result = await api("ownerStatus", {}, true); state.ownerOrder = result.order;
}
function clearOwnerAccess() {
  pinTarget = null;
  privateCardLink = ""; $("privateAccessLink").value = "";
  $("paymentDialog").close();
  $("developerPassword").value = "";
  $("pinInput").value = ""; $("setupPin").value = ""; $("setupPinConfirm").value = ""; $("pinSetupDialog").close();
  try { sessionStorage.removeItem("rifapro.cardAccess"); } catch { /* Optional storage. */ }
}
async function enterOwnerLink(link) {
  const url = new URL(link), key = new URLSearchParams(url.hash.slice(1)).get("acceso"), id = url.searchParams.get("rifa");
  if (!id || !/^[a-f0-9-]{72}$/.test(key || "")) throw new Error("Pega el enlace privado completo que guardaste al crear tu cartón.");
  if (hasSession()) logout(false);
  const result = await api("ownerLogin", { rifaId: id, ownerKey: key });
  privateCardLink = cardOwnerLink(id, key);
  try { sessionStorage.setItem("rifapro.cardAccess", privateCardLink); } catch { /* Optional recovery. */ }
  useOwnerSession(result);
  if (result.order.status === "active") { state.entryLoading = true; render(); refresh(); } else fillPayment(result.order);
}
async function copyText(value, message) {
  try {
    await navigator.clipboard.writeText(value); toast(message);
    const key = value === PAYMENT_KEY;
    $("toast").classList.add("clipboard-toast");
    $("toast").innerHTML = '<span class="clipboard-check" aria-hidden="true">✓</span><span><strong>' + (key ? 'Llave copiada' : 'Enlace copiado') + '</strong><small>' + escapeHTML(key ? '320 875 9798 · Lista para pegar en tu app.' : 'Listo para pegar y compartir.') + '</small></span>';
  }
  catch { notice("Copia estos datos", value); }
}
function downloadText(name, text, type = "text/plain;charset=utf-8") {
  const url = URL.createObjectURL(new Blob([text], { type })), a = document.createElement("a"); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
}
async function loadOrders(event) {
  if (!isPlatform() || ordersLoading) return;
  const token = state.token; ordersLoading = true; if (event?.type === "click") $("refreshOrders").disabled = true;
  try {
    const result = await api("developerDashboard", {}, true); if (state.token !== token) return;
    state.capabilities = result.capabilities || state.capabilities;
    const cards = result.cards;
    const pending = cards.filter((o) => ["pending", "reviewing"].includes(o.status) && !o.expired);
    const active = cards.filter((o) => o.status === "active" && !o.expired);
    markup("developerStats", [["Cartones", cards.length],["Por aprobar",pending.length],["Activos",active.length],["Vencidos",cards.filter((o) => o.expired && o.status !== "deleted").length]].map(([label,count]) => `<article class="kpi"><strong>${label}</strong>${count}</article>`).join(""));
    const filter = $("developerFilter").value;
    const visible = cards.filter((o) => filter === "all" || (filter === "pending" ? pending.includes(o) : filter === "active" ? active.includes(o) : o.expired && o.status !== "deleted"));
    markup("ordersList", visible.length ? visible.map((o) => `<article class="order-row"><div><strong>${escapeHTML(o.titulo)}</strong><small>${escapeHTML(o.contact || "Rifa anterior")} · <span class="card-status ${o.status}">${escapeHTML(o.status === "deleted" ? "Eliminado" : o.expired ? "Vencido" : paymentLabels[o.status])}</span></small><small>Sorteo: ${escapeHTML(formatDate(o.fecha))}${o.legacy ? " · Cartón anterior" : " · " + currencyText(o.amount) + " COP"}</small><small class="order-id">${escapeHTML(o.reference || (o.legacy ? "Cartón anterior" : cardReference(o)))}</small>${o.pinConflict ? '<small class="form-error">PIN repetido en una versión anterior. Asigna uno nuevo.</small>' : ''}</div><div class="order-actions">${pending.includes(o) ? `<button type="button" class="button button-outline button-small" data-approve-card="${escapeHTML(o.rifaId)}" data-card-amount="${Number(o.amount)}">Aprobar ${currencyText(o.amount)}</button>` : o.status === "active" ? `<a class="button button-outline button-small" target="_blank" rel="noopener noreferrer" href="${escapeHTML(rafflePublicLink(o.rifaId).href)}">Ver rifa ↗</a>` : ""}${o.status !== "deleted" ? `<button type="button" class="button button-quiet button-small" data-reset-pin="${escapeHTML(o.rifaId)}">Cambiar PIN</button>` : ''}</div></article>`).join("") : '<p class="muted">No hay cartones en este estado.</p>');
  } catch (error) { status("No se pudo actualizar · Conservamos la vista anterior", "error"); if (event?.type === "click") toast(error.message); }
  finally { ordersLoading = false; $("refreshOrders").disabled = false; }
}
function initCartones() {
  $("pinSetupButton").addEventListener("click", () => { pinTarget = state.activeId; $("pinSetupError").textContent = ""; $("pinSetupForm").reset(); $("pinSetupDialog").showModal(); });
  $("pinSetupForm").addEventListener("submit", async (event) => {
    event.preventDefault(); const pin = $("setupPin").value;
    if (!/^\d{4}$/.test(pin) || pin !== $("setupPinConfirm").value) { $("pinSetupError").textContent = "Los PIN deben coincidir y tener cuatro dígitos."; return; }
    $("savePinButton").disabled = true;
    try { const result = await api("setOwnerPin", { rifaId: pinTarget, pin }, true); if (!isPlatform()) state.ownerOrder = result.order; $("pinSetupDialog").close(); $("setupPin").value = ""; $("setupPinConfirm").value = ""; render(); if (isPlatform()) loadOrders(); toast("PIN actualizado. Ya se puede entrar al cartón con sus cuatro dígitos."); }
    catch (error) { $("pinSetupError").textContent = error.message; }
    finally { $("savePinButton").disabled = false; }
  });
  $("developerForm").addEventListener("submit", async (event) => {
    event.preventDefault(); $("developerLoginButton").disabled = true; $("developerError").textContent = "";
    const generation = state.generation;
    try {
      const result = await api("developerLogin", { password: $("developerPassword").value });
      if (generation !== state.generation) return;
      if (result.role !== "developer" || !result.token) throw new Error("No se confirmó el acceso del desarrollador.");
      state.token = result.token; state.expiresAt = result.expiresAt; state.role = "developer"; state.ownerOrder = null; state.generation++;
      render(); await loadOrders();
    } catch (error) { $("developerError").textContent = error.message; }
    finally { $("developerPassword").value = ""; $("developerLoginButton").disabled = false; }
  });
  $("developerFilter").addEventListener("change", loadOrders);
  $("chooseNumbers").addEventListener("click", scrollToNumbers);
  $("publicCreateButton").addEventListener("click", startPaidCard);
  $("ownerAccessButton").addEventListener("click", () => openPinLogin(true));
  $("copyPaymentKey").addEventListener("click", () => copyText(PAYMENT_KEY, "Llave copiada: 320 875 9798"));
  $("copyPrivateLink").addEventListener("click", () => copyText(currentCardLink(), state.ownerOrder?.hasPin ? "Enlace del cartón copiado." : "Acceso anterior copiado. Guárdalo solo para ti."));
  $("downloadPrivateLink").addEventListener("click", () => downloadText("mi-carton-rifa-pro.txt", state.ownerOrder?.hasPin ? "RIFA PRO · MI CARTÓN\nAbre este enlace, pulsa Administrador y escribe tu PIN de 4 dígitos.\nLos participantes pueden abrir el enlace sin iniciar sesión.\n\n" + currentCardLink() : "RIFA PRO · ACCESO PRIVADO ANTERIOR\nNo compartas este enlace con participantes.\n\n" + privateCardLink));
  $("downloadPayment").addEventListener("click", showPaymentQr);
  $("paidWhatsApp").addEventListener("click", () => { const token = state.token; api("reportPayment", {}, true).then((r) => { if (state.role === "owner" && state.token === token) { state.ownerOrder = r.order; $("cardPaymentStatus").textContent = paymentLabels[r.order.status]; renderCartones(); } }).catch((error) => toast(error.message)); });
  $("enterMyCard").addEventListener("click", () => { $("paymentDialog").close(); refresh(); requestAnimationFrame(scrollToNumbers); });
  $("openPayment").addEventListener("click", () => fillPayment(state.ownerOrder));
  $("refreshOrders").addEventListener("click", loadOrders);
  $("ordersList").addEventListener("click", async (event) => {
    const reset = event.target.closest("[data-reset-pin]");
    if (reset && isPlatform()) { pinTarget = reset.dataset.resetPin; $("pinSetupError").textContent = ""; $("pinSetupForm").reset(); $("pinSetupDialog").showModal(); return; }
    const button = event.target.closest("[data-approve-card]"); if (!button || button.disabled) return;
    if (!await confirmAction("¿Confirmas que recibiste " + currencyText(button.dataset.cardAmount) + "?", "Aprueba solo después de verificar la transferencia y el comprobante de esta referencia. Se activará el enlace público del cartón.", "Verificado · Activar cartón")) return;
    button.disabled = true; button.textContent = "Activando…";
    try { await api("approveCard", { rifaId: button.dataset.approveCard }, true); toast("Cartón activado."); await loadOrders(); refresh(); } catch (error) { toast(error.message); button.disabled = false; button.textContent = "Reintentar activación"; }
  });
  const incoming = new URL(location.href);
  if (!developerPage && incoming.hash.startsWith("#acceso=")) {
    const privateUrl = incoming.href; history.replaceState(null, "", incoming.pathname + incoming.search);
    enterOwnerLink(privateUrl).catch((error) => notice("No pudimos abrir tu cartón", error.message));
  }
  renderCartones();
}
document.addEventListener("DOMContentLoaded", initCartones);
