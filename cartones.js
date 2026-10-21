/* Independent cards: no payment is considered verified by a client-side click. */
"use strict";
const PAYMENT_KEY = "3208759798";
const RECEIPT_PHONE = "573246394689";
let cardDraft = null;
let privateCardLink = "";
let ordersLoading = false;
const paymentLabels = { pending: "Pendiente de pago", reviewing: "Comprobante por revisar", active: "Cartón activo", deleted: "Cartón eliminado" };

function scrollToNumbers() {
  $("numeros").scrollIntoView({ block: "start", behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "instant" : "smooth" });
  $("numbersTitle").tabIndex = -1; $("numbersTitle").focus({ preventScroll: true });
}
function renderCartones() {
  $("newRaffleButton").hidden = !isPlatform();
  $("adminButtonText").textContent = hasSession() ? "Cerrar sesión" : "Administrador";
  $("publicCreateButton").hidden = hasSession();
  document.querySelector(".creator-entry").hidden = hasSession();
  $("ownerCard").hidden = !(hasSession() && state.role === "owner");
  if (state.ownerOrder) {
    $("ownerCardTitle").textContent = state.ownerOrder.titulo;
    $("ownerCardStatus").textContent = paymentLabels[state.ownerOrder.status] || "Por confirmar";
    $("cardPaymentStatus").textContent = paymentLabels[state.ownerOrder.status] || "Por confirmar";
    $("paidWhatsApp").hidden = ["active", "deleted"].includes(state.ownerOrder.status);
    $("enterMyCard").hidden = state.ownerOrder.status !== "active";
  }
  $("ordersPanel").hidden = !(isPlatform() && state.capabilities.independentRaffles);
  if ($( "ordersPanel").hidden) $("ordersList").replaceChildren();
  const r = activeRaffle();
  $("expiryNotice").hidden = !r?.expiresAt;
  if (r?.expiresAt) $("expiryNotice").textContent = raffleExpired(r) ? "Cartón vencido · Los registros se conservan para consulta." : "Disponible hasta el " + new Intl.DateTimeFormat("es-CO", { day: "numeric", month: "long", year: "numeric", timeZone: "America/Bogota" }).format(new Date(Date.parse(r.expiresAt) - 1)) + " · Dos días después del sorteo.";
  $("chooseNumbers").hidden = !r;
  $("chooseNumbers").disabled = raffleExpired(r);
  if (hasSession() && state.role === "owner" && !isAdmin()) status("Cartón pendiente de activación", "online");
}
function startPaidCard() {
  if (hasSession()) { toast("Cierra tu sesión antes de crear otro cartón independiente."); return; }
  if (!state.capabilities.independentRaffles) { notice("Estamos habilitando los nuevos cartones", "El administrador debe aplicar la versión actualizada de Code.gs para crear cartones independientes."); return; }
  $("raffleForm").reset(); $("raffleForm").dataset.mode = "paid";
  $("raffleDialogTitle").textContent = "Tu rifa, tu propio espacio.";
  $("creatorFields").hidden = false; $("creatorPhone").required = true;
  $("createRaffleButton").textContent = "Continuar al pago · $1.000 →";
  $("raffleError").textContent = "";
  if (!cardDraft) cardDraft = { rifaId: uuid(), ownerKey: uuid() + uuid() };
  $("raffleDialog").showModal();
}
function rafflePublicLink(id) { const url = new URL(location.href); url.search = ""; url.hash = ""; url.searchParams.set("rifa", id); return url; }
function cardOwnerLink(id, key) { const url = rafflePublicLink(id); url.hash = "acceso=" + key; return url.href; }
function fillPayment(order) {
  $("paymentReference").textContent = order.rifaId;
  $("cardPaymentStatus").textContent = paymentLabels[order.status] || "Por confirmar";
  $("privateAccessLink").value = privateCardLink;
  $("privateAccessSection").hidden = !privateCardLink;
  $("paidWhatsApp").href = "https://wa.me/" + RECEIPT_PHONE + "?text=" + encodeURIComponent("Hola, te escribo porque ya pagué mi cartón de rifa virtual. Te dejo el comprobante aquí:\n\nCartón: " + order.titulo + "\nReferencia: " + order.rifaId + "\nValor: $1.000 COP.\nAdjuntaré el comprobante en este chat.");
  $("paidWhatsApp").hidden = order.status === "active" || order.status === "deleted";
  $("enterMyCard").hidden = order.status !== "active";
  if (!$("paymentDialog").open) $("paymentDialog").showModal();
}
async function createPaidCard() {
  const button = $("createRaffleButton"); if (button.disabled) return;
  const payload = { ...cardDraft, contact: $("creatorPhone").value.trim(), ...Object.fromEntries(["titulo", "precio", "cantidad", "premio", "loteria", "fecha", "hora", "responsable", "condiciones"].map((id) => [id, $(id).value.trim()])) };
  // Keep this recovery link before sending: an uncertain result can be retried safely.
  privateCardLink = cardOwnerLink(payload.rifaId, payload.ownerKey);
  try { sessionStorage.setItem("rifapro.cardAccess", privateCardLink); } catch { /* User can download the link. */ }
  button.disabled = true; button.textContent = "Preparando tu cartón…"; $("raffleError").textContent = "";
  try {
    const result = await api("orderCard", payload);
    const session = await api("ownerLogin", { rifaId: payload.rifaId, ownerKey: payload.ownerKey });
    useOwnerSession(session); $("raffleDialog").close(); fillPayment(result.order); cardDraft = null;
  } catch (error) { $("raffleError").textContent = error.message + " Puedes reintentar sin crear otro cartón."; }
  finally { button.disabled = false; button.textContent = "Continuar al pago · $1.000 →"; }
}
function useOwnerSession(result) {
  state.token = result.token; state.expiresAt = result.expiresAt; state.role = "owner"; state.ownerOrder = result.order;
  state.activeId = result.rifaId; state.rifas = []; state.records = []; state.summary = null;
  state.generation++; state.online = true; state.capabilities.independentRaffles = true; state.capabilities.snapshot = true;
  render();
}
async function refreshOwnerOrder() {
  if (state.ownerOrder?.status === "active") return;
  const result = await api("ownerStatus", {}, true); state.ownerOrder = result.order;
}
function clearOwnerAccess() {
  privateCardLink = ""; $("privateAccessLink").value = ""; $("ownerLinkInput").value = "";
  $("paymentDialog").close(); $("ownerLoginDialog").close();
  try { sessionStorage.removeItem("rifapro.cardAccess"); } catch { /* Optional storage. */ }
}
async function enterOwnerLink(link) {
  const url = new URL(link), key = new URLSearchParams(url.hash.slice(1)).get("acceso"), id = url.searchParams.get("rifa");
  if (!id || !/^[a-f0-9-]{72}$/.test(key || "")) throw new Error("Pega el enlace privado completo que guardaste al crear tu cartón.");
  if (hasSession()) logout(false);
  const result = await api("ownerLogin", { rifaId: id, ownerKey: key });
  privateCardLink = cardOwnerLink(id, key);
  try { sessionStorage.setItem("rifapro.cardAccess", privateCardLink); } catch { /* Optional recovery. */ }
  useOwnerSession(result); $("ownerLoginDialog").close();
  if (result.order.status === "active") refresh(); else fillPayment(result.order);
}
async function copyText(value, message) {
  try { await navigator.clipboard.writeText(value); toast(message); }
  catch { notice("Copia estos datos", value); }
}
function downloadText(name, text, type = "text/plain;charset=utf-8") {
  const url = URL.createObjectURL(new Blob([text], { type })), a = document.createElement("a"); a.href = url; a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function downloadPaymentCard() {
  // A native instruction card. No QR is generated or substituted for a bank QR.
  const canvas = document.createElement("canvas"); canvas.width = 1080; canvas.height = 1200;
  const ctx = canvas.getContext("2d"); ctx.fillStyle = "#f5f2eb"; ctx.fillRect(0, 0, 1080, 1200);
  ctx.fillStyle = "#17283d"; ctx.fillRect(60, 60, 960, 1080); ctx.textAlign = "center";
  ctx.fillStyle = "#ded5c3"; ctx.font = "bold 30px -apple-system, sans-serif"; ctx.fillText("RIFA PRO · TU CARTÓN VIRTUAL", 540, 170);
  ctx.fillStyle = "#ffffff"; ctx.font = "bold 100px -apple-system, sans-serif"; ctx.fillText("$1.000 COP", 540, 340);
  ctx.font = "bold 52px -apple-system, sans-serif"; ctx.fillText("Nequi / Bre-B", 540, 480);
  ctx.font = "30px -apple-system, sans-serif"; ctx.fillText("LLAVE / CELULAR", 540, 570);
  ctx.font = "bold 70px -apple-system, sans-serif"; ctx.fillText("320 875 9798", 540, 665);
  ctx.font = "42px -apple-system, sans-serif"; ctx.fillText("Elkin Rangel", 540, 755);
  ctx.font = "26px -apple-system, sans-serif"; ctx.fillText("Envía el comprobante por WhatsApp", 540, 880); ctx.fillText("324 639 4689", 540, 925);
  ctx.font = "22px -apple-system, sans-serif"; ctx.fillText("Referencia: " + (state.ownerOrder?.rifaId || ""), 540, 1020, 880);
  ctx.fillStyle = "#ded5c3"; ctx.fillText("La activación requiere verificar el pago.", 540, 1080);
  const a = document.createElement("a"); a.download = "pago-carton-rifa-pro.png"; a.href = canvas.toDataURL("image/png"); a.click();
}
async function loadOrders() {
  if (!isPlatform() || !state.capabilities.independentRaffles || ordersLoading) return;
  const token = state.token; ordersLoading = true; $("refreshOrders").disabled = true;
  try {
    const result = await api("listOrders", {}, true); if (state.token !== token) return;
    const pending = result.orders.filter((o) => ["pending", "reviewing"].includes(o.status));
    $("ordersList").innerHTML = pending.length ? pending.map((o) => `<article class="order-row"><div><strong>${escapeHTML(o.titulo)}</strong><small>${escapeHTML(o.contact)} · ${escapeHTML(paymentLabels[o.status])}</small><small class="order-id">${escapeHTML(o.rifaId)}</small></div><button type="button" class="button button-outline button-small" data-approve-card="${escapeHTML(o.rifaId)}">Aprobar $1.000</button></article>`).join("") : '<p class="muted">No hay cartones pendientes de aprobación.</p>';
  } catch (error) { toast(error.message); }
  finally { ordersLoading = false; $("refreshOrders").disabled = false; }
}
function initCartones() {
  $("chooseNumbers").addEventListener("click", scrollToNumbers);
  $("publicCreateButton").addEventListener("click", startPaidCard);
  $("ownerAccessButton").addEventListener("click", () => { $("ownerLoginError").textContent = ""; try { $("ownerLinkInput").value = sessionStorage.getItem("rifapro.cardAccess") || ""; } catch {} $("ownerLoginDialog").showModal(); });
  $("ownerLoginForm").addEventListener("submit", async (event) => { event.preventDefault(); $("ownerLoginButton").disabled = true; try { await enterOwnerLink($("ownerLinkInput").value.trim()); } catch (error) { $("ownerLoginError").textContent = error.message; } finally { $("ownerLoginButton").disabled = false; } });
  $("copyPaymentKey").addEventListener("click", () => copyText(PAYMENT_KEY, "Llave copiada: 320 875 9798"));
  $("copyPrivateLink").addEventListener("click", () => copyText(privateCardLink, "Enlace privado copiado. Guárdalo solo para ti."));
  $("downloadPrivateLink").addEventListener("click", () => downloadText("mi-carton-acceso-privado.txt", "RIFA PRO · ACCESO PRIVADO\nNo compartas este enlace con participantes.\n\n" + privateCardLink));
  $("downloadPayment").addEventListener("click", downloadPaymentCard);
  $("paidWhatsApp").addEventListener("click", () => { const token = state.token; api("reportPayment", {}, true).then((r) => { if (state.role === "owner" && state.token === token) { state.ownerOrder = r.order; $("cardPaymentStatus").textContent = paymentLabels[r.order.status]; renderCartones(); } }).catch((error) => toast(error.message)); });
  $("enterMyCard").addEventListener("click", () => { $("paymentDialog").close(); refresh(); scrollToNumbers(); });
  $("openPayment").addEventListener("click", () => fillPayment(state.ownerOrder));
  $("refreshOrders").addEventListener("click", loadOrders);
  $("ordersList").addEventListener("click", async (event) => {
    const button = event.target.closest("[data-approve-card]"); if (!button || button.disabled) return;
    if (!await confirmAction("¿Confirmas que recibiste $1.000?", "Aprueba solo después de verificar la transferencia y el comprobante de esta referencia. Se activará el enlace público del cartón.", "Verificado · Activar cartón")) return;
    button.disabled = true; button.textContent = "Activando…";
    try { await api("approveCard", { rifaId: button.dataset.approveCard }, true); toast("Cartón activado."); await loadOrders(); refresh(); } catch (error) { toast(error.message); button.disabled = false; button.textContent = "Reintentar activación"; }
  });
  $("adminDialog").addEventListener("close", () => { if (isPlatform()) loadOrders(); });
  const incoming = new URL(location.href);
  if (incoming.hash.startsWith("#acceso=")) {
    const privateUrl = incoming.href; history.replaceState(null, "", incoming.pathname + incoming.search);
    enterOwnerLink(privateUrl).catch((error) => notice("No pudimos abrir tu cartón", error.message));
  }
  renderCartones();
}
document.addEventListener("DOMContentLoaded", initCartones);
