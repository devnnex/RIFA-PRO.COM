/* Mobile creation, readable currency and native image saving. No reservation data is cached. */
"use strict";
let creationStep = 1;
let qrAsset = null;
let qrAssetReady = null;
function amountValue(value) { return Number(String(value).replace(/\D/g, "")) || 0; }
function currencyText(value) { return "$" + new Intl.NumberFormat("es-CO", { maximumFractionDigits: 0 }).format(amountValue(value)); }
function prizeText(value) {
  const text = String(value || "").trim();
  return /^\$?\s*\d[\d.,\s]*$/.test(text) ? currencyText(text) : text.replace(/\$\s*\d[\d.,]*/g, currencyText);
}
function formatMoneyField(input) {
  const raw = input.value, caret = input.selectionStart, mode = input.dataset.money;
  const digitsBefore = (raw.slice(0, caret ?? raw.length).match(/\d/g) || []).length;
  let next = raw;
  if (mode === "amount") next = /\d/.test(raw) ? currencyText(raw) : "";
  else if (/^\s*\$?\s*\d[\d.,\s]*$/.test(raw)) next = currencyText(raw) + (raw.endsWith(" ") ? " " : "");
  else next = raw.replace(/\$\s*\d[\d.,]*/g, currencyText);
  if (next === raw) return;
  input.value = next;
  if (caret != null && document.activeElement === input) {
    let position = 0, count = 0;
    if (mode === "amount" || /^\s*\$?\s*\d[\d.,\s]*$/.test(raw)) {
      while (position < next.length && count < digitsBefore) { if (/\d/.test(next[position])) count++; position++; }
    } else position = Math.max(0, caret + next.length - raw.length);
    input.setSelectionRange(position, position);
  }
}
function updateCreationPreview() {
  $("creationPreviewTitle").textContent = $("titulo").value.trim() || "Tu próxima gran rifa";
  $("creationPreviewPrice").textContent = ($("cantidad").value || "100") + " números · " + currencyText($("precio").value) + " cada uno";
}
function validateCreationStep(step) {
  const section = $(step === 1 ? "createStepOne" : "createStepTwo");
  const price = amountValue($("precio").value);
  $("precio").setCustomValidity(price >= 1 && price <= 100000000 && Number.isSafeInteger(price) ? "" : "Escribe un valor entre $1 y $100.000.000.");
  $("creatorPinConfirm").setCustomValidity($("creatorPinConfirm").value === $("creatorPin").value ? "" : "Los dos PIN deben coincidir.");
  const invalid = [...section.querySelectorAll("input")].find((input) => !input.checkValidity());
  if (!invalid) return true;
  if (invalid.closest("details")) invalid.closest("details").open = true;
  setCreationStep(step); invalid.reportValidity(); return false;
}
function validateCreation() { return validateCreationStep(1) && validateCreationStep(2); }
function setCreationStep(step) {
  creationStep = step;
  $("createStepOne").hidden = step !== 1; $("createStepTwo").hidden = step !== 2;
  $("nextCreateStep").hidden = step !== 1; $("previousCreateStep").hidden = step !== 2; $("createRaffleButton").hidden = step !== 2;
  for (const [number, id] of [[1, "stepOneLabel"], [2, "stepTwoLabel"]]) {
    if (step === number) $(id).setAttribute("aria-current", "step"); else $(id).removeAttribute("aria-current");
  }
  $("stepOneLabel").classList.toggle("step-done", step === 2);
  $("raffleDialog").querySelector(".create-body").scrollTop = 0;
  $("raffleError").textContent = "";
}
function resetCreation() { setCreationStep(1); formatMoneyField($("precio")); updateCreationPreview(); }
function renderExperience() {
  const linked = new URL(location.href).searchParams.has("rifa"), visitor = linked && !hasSession() && !developerPage;
  document.documentElement.classList.toggle("shared-link", linked);
  document.body.classList.toggle("participant-view", visitor);
  $("developerAccessButton").hidden = linked;
  $("adminButton").hidden = visitor || (developerPage && !hasSession());
  $("acquireCardButton").hidden = !visitor;
  document.querySelector(".page-intro").hidden = developerPage || linked;
  document.querySelector(".creator-entry").hidden = hasSession() || linked;
  $("ownerCard").hidden ||= visitor;
  renderWinner();
}
const seenWinners = new Set();
let winnerDraft = null, publishingWinner = false;
function renderWinner() {
  const r = activeRaffle(), winner = r?.winner;
  $("publishWinnerButton").hidden = !isAdmin();
  $("publishWinnerButton").textContent = winner ? "✦ Actualizar ganador" : "✦ Publicar ganador";
  $("publishedWinner").hidden = !isAdmin() || !winner;
  if (winner) $("publishedWinner").textContent = "Ganador publicado · " + winner.numeroId + " · " + winner.nombre;
  if (!winner || hasSession() || developerPage || state.entryLoading || !state.online || !state.protocol || document.hidden) return;
  const key = r.id + ":" + winner.revision + ":" + winner.publishedAt;
  if (seenWinners.has(key) || document.querySelector("dialog[open]")) return;
  seenWinners.add(key);
  $("winningNumber").textContent = winner.numeroId;
  $("winnerName").textContent = winner.nombre;
  $("winnerConfetti").replaceChildren();
  if (!matchMedia("(prefers-reduced-motion: reduce)").matches) {
    for (let i = 0; i < 36; i++) {
      const piece = document.createElement("i");
      piece.style.cssText = `--x:${(i * 29) % 100}%;--delay:${(i % 9) * .07}s;--turn:${i % 2 ? 240 : -210}deg;--color:${["#d4ad50", "#32926a", "#9bd5bc", "#c6a9dd"][i % 4]}`;
      $("winnerConfetti").appendChild(piece);
    }
  }
  $("winnerDialog").showModal();
}
function previewWinner() {
  const number = $("winnerNumber").value.trim();
  const record = /^\d{1,6}$/.test(number) ? state.records.find((n) => Number(n.numeroId) === Number(number) && n.estado !== "libre") : null;
  $("winnerNamePreview").textContent = record?.clienteNombre || (number ? "Este número aún no tiene participante" : "Escribe el número ganador");
  $("winnerNumber").setCustomValidity(record?.clienteNombre ? "" : "Elige un número asignado a un participante.");
}
function openWinnerEditor() {
  if (!isAdmin() || !activeRaffle()) return;
  if (!state.capabilities.publishWinner) { notice("La publicación estará disponible pronto", "Vuelve a intentarlo cuando el organizador haya habilitado esta función."); return; }
  winnerDraft = { rifaId: state.activeId, revision: activeRaffle().winner?.revision || 0, requestId: uuid() };
  $("winnerForm").reset(); $("winnerError").textContent = "";
  $("winnerNumber").value = activeRaffle().winner?.numeroId || ""; previewWinner();
  $("winnerEditorDialog").showModal();
}
async function saveWinner(event) {
  event.preventDefault();
  if (publishingWinner || !isAdmin() || !winnerDraft || winnerDraft.rifaId !== state.activeId) return;
  previewWinner(); if (!$("winnerForm").reportValidity()) return;
  const token = state.token, draft = { ...winnerDraft, numeroId: $("winnerNumber").value.trim() };
  publishingWinner = true; $("saveWinnerButton").disabled = true; $("winnerNumber").readOnly = true;
  $("winnerEditorDialog").querySelector("[data-close]").disabled = true;
  $("saveWinnerButton").textContent = "Publicando…"; $("winnerError").textContent = "";
  try {
    const result = await api("publishWinner", draft, true);
    if (token !== state.token || state.activeId !== draft.rifaId) return;
    if (!result.ok || !result.winner) throw new Error("No se confirmó la publicación. Intenta de nuevo.");
    state.writeEpoch++; activeRaffle().winner = result.winner;
    $("winnerEditorDialog").close(); render(); toast("Ganador publicado. Ya aparece en el enlace de tu rifa.");
  } catch (error) {
    $("winnerError").textContent = error.message;
    if (error.code === "STALE") { await refresh(); winnerDraft.revision = activeRaffle()?.winner?.revision || 0; winnerDraft.requestId = uuid(); }
  } finally {
    publishingWinner = false; $("saveWinnerButton").disabled = false; $("winnerNumber").readOnly = false;
    $("winnerEditorDialog").querySelector("[data-close]").disabled = false;
    $("saveWinnerButton").textContent = "Confirmar y publicar ganador ✦";
  }
}
function installDevice() {
  const ua = navigator.userAgent;
  const ios = /iPad|iPhone|iPod/.test(ua) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  return { ios, android: /Android/.test(ua), chrome: /Chrome|CriOS|Edg/.test(ua) };
}
function selectInstallBrowser(browser) {
  $("safariInstall").hidden = browser !== "safari"; $("chromeInstall").hidden = browser !== "chrome";
  document.querySelectorAll("[data-install-browser]").forEach((button) => button.setAttribute("aria-pressed", String(button.dataset.installBrowser === browser)));
}
function showInstallGuide() {
  const device = installDevice();
  const steps = device.ios ? ["Abre esta página en <strong>Chrome</strong>.", "A la derecha de la barra de direcciones, toca <strong>Compartir</strong>.", "Elige <strong>Añadir a la pantalla de inicio</strong>.", "Confirma el nombre y toca <strong>Añadir</strong>."] : device.android ? ["Abre esta página en <strong>Chrome</strong>.", "Toca <strong>Más (⋮)</strong>, a la derecha de la barra de direcciones.", "Elige <strong>Instalar y crear acceso directo → Instalar</strong>. En algunas versiones aparece <strong>Añadir a la pantalla de inicio → Instalar</strong>.", "Confirma <strong>Instalar</strong> y abre Rifa Pro desde su icono."] : ["Abre esta página en <strong>Chrome</strong>.", "Arriba a la derecha, toca <strong>Más (⋮)</strong>.", "Elige <strong>Enviar, guardar y compartir → Instalar página como aplicación</strong>.", "Confirma <strong>Instalar</strong>. También puedes usar el icono Instalar de la barra de direcciones, si aparece."];
  $("chromeDeviceTitle").textContent = "Chrome · " + (device.ios ? "iPhone / iPad" : device.android ? "Android" : "Computador");
  $("chromeInstallSteps").innerHTML = steps.map((step) => "<li>" + step + "</li>").join("");
  selectInstallBrowser(device.ios && !device.chrome ? "safari" : "chrome");
  $("nativeInstallButton").hidden = !state.installPrompt;
  $("installSupport").href = "https://wa.me/" + RECEIPT_PHONE + "?text=" + encodeURIComponent("Hola, equipo de Rifa Pro 👋 Intenté agregar la app a mi pantalla de inicio, pero no pude completar el proceso.");
  $("installDialog").showModal();
}
let appUpdateReady = false, lastAppInteraction = Date.now();
function canApplyAppUpdate() {
  return !hasSession() && !state.busy && !publishingWinner && !state.pending.size && !state.pendingRaffles.size && !cardDraft &&
    !document.querySelector("dialog[open]") && !document.activeElement?.matches("input,textarea,select") && !$("numberSearch").value;
}
function applyAppUpdate() {
  if (!appUpdateReady || state.busy || publishingWinner || state.pending.size || state.pendingRaffles.size || document.querySelector("dialog[open]")) {
    toast("Termina lo que estás haciendo y luego toca Actualizar. Tus cambios se conservarán."); return;
  }
  location.reload();
}
function setupAppUpdates() {
  if (!("serviceWorker" in navigator)) return;
  const wasControlled = Boolean(navigator.serviceWorker.controller);
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (!wasControlled) return; // The first installation does not interrupt the page.
    appUpdateReady = true; $("appUpdateNotice").hidden = false;
    if (canApplyAppUpdate() && Date.now() - lastAppInteraction > 15000) applyAppUpdate();
  });
  $("applyAppUpdate").addEventListener("click", applyAppUpdate);
  for (const name of ["pointerdown", "keydown", "input"]) document.addEventListener(name, () => { lastAppInteraction = Date.now(); }, { passive: true });
  navigator.serviceWorker.register("./sw.js", { updateViaCache: "none" }).then((registration) => {
    let lastCheck = 0;
    const check = () => {
      if (!navigator.onLine || document.hidden || Date.now() - lastCheck < 60000) return;
      lastCheck = Date.now(); registration.update().catch(() => {});
    };
    check(); window.addEventListener("online", check);
    document.addEventListener("visibilitychange", () => { if (!document.hidden) check(); });
    setInterval(check, 5 * 60000);
  }).catch(() => {});
  setInterval(() => { if (appUpdateReady && !document.hidden && canApplyAppUpdate() && Date.now() - lastAppInteraction > 30000) applyAppUpdate(); }, 10000);
}
function preloadPaymentQr() {
  const source = CONFIG.paymentQrUrl;
  if (!source) return Promise.resolve(null);
  const img = $("paymentQrImage"); img.src = source;
  $("qrOpenOriginal").href = source;
  return Promise.all([img.decode().catch(() => {}), fetch(source, { cache: "force-cache" }).then((r) => { if (!r.ok) throw new Error(); return r.blob(); })])
    .then(([, blob]) => { if (!img.naturalWidth || !blob.type.startsWith("image/")) return null; qrAsset = new File([blob], "qr-elkin-rangel.png", { type: blob.type }); return qrAsset; }).catch(() => null);
}
async function showPaymentQr() {
  $("qrDialog").showModal();
  // Precaching starts at page entry; the viewer never generates or redraws bank codes.
  const asset = qrAsset || await qrAssetReady;
  $("qrImageFrame").hidden = !asset; $("saveQrImage").hidden = !asset;
  $("qrUnavailable").hidden = Boolean(asset); $("qrSaveHint").hidden = !asset;
}
async function savePaymentQr() {
  if (!qrAsset) return;
  try {
    if (navigator.canShare?.({ files: [qrAsset] })) {
      await navigator.share({ files: [qrAsset], title: "QR de pago · Elkin Rangel" });
      return;
    }
    const a = document.createElement("a"); a.href = $("paymentQrImage").src; a.download = "qr-elkin-rangel.png"; a.click();
    $("qrSaveHint").textContent = "La imagen sigue abierta aquí. En iPhone, mantenla pulsada y elige Guardar en Fotos; también puedes usar Compartir.";
    toast("Imagen lista. Usa las opciones de tu teléfono para guardarla.");
  } catch (error) { if (error.name !== "AbortError") toast("Mantén pulsada la imagen para guardarla desde las opciones de tu teléfono."); }
}
function setupModalScrolling() {
  let locked = false, savedY = 0;
  function syncLock() {
    const open = Boolean(document.querySelector("dialog[open]"));
    if (open === locked) return;
    locked = open;
    if (open) { savedY = window.scrollY; document.body.style.top = -savedY + "px"; document.body.classList.add("modal-locked"); }
    else { document.body.classList.remove("modal-locked"); document.body.style.top = ""; window.scrollTo({ top: savedY, behavior: "instant" }); }
  }
  const observer = new MutationObserver(syncLock);
  document.querySelectorAll("dialog").forEach((dialog) => {
    observer.observe(dialog, { attributes: true, attributeFilter: ["open"] });
    dialog.addEventListener("close", syncLock);
    let lastX = 0, lastY = 0;
    dialog.addEventListener("touchstart", (event) => { if (event.touches.length === 1) { lastX = event.touches[0].clientX; lastY = event.touches[0].clientY; } }, { passive: true });
    dialog.addEventListener("touchmove", (event) => {
      if (event.touches.length !== 1) return;
      const point = event.touches[0], dy = lastY - point.clientY, dx = lastX - point.clientX;
      lastY = point.clientY; lastX = point.clientX;
      if (Math.abs(dx) > Math.abs(dy)) return;
      const scroller = dialog.querySelector(".create-body") || dialog;
      if (!scroller.contains(event.target)) { scroller.scrollTop += dy; event.preventDefault(); return; }
      const atTop = scroller.scrollTop <= 0, atEnd = scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 1;
      if ((dy < 0 && atTop) || (dy > 0 && atEnd)) event.preventDefault();
    }, { passive: false });
  });
}
function initExperience() {
  $("publishWinnerButton").addEventListener("click", openWinnerEditor);
  $("winnerForm").addEventListener("submit", saveWinner);
  $("winnerNumber").addEventListener("input", () => { if (winnerDraft) winnerDraft.requestId = uuid(); previewWinner(); });
  $("winnerEditorDialog").addEventListener("cancel", (event) => { if (publishingWinner) event.preventDefault(); });
  document.querySelectorAll("dialog").forEach((dialog) => dialog.addEventListener("close", () => setTimeout(renderWinner, 0)));
  document.querySelectorAll("[data-install-browser]").forEach((button) => button.addEventListener("click", () => selectInstallBrowser(button.dataset.installBrowser)));
  $("nativeInstallButton").addEventListener("click", async () => {
    const prompt = state.installPrompt; if (!prompt) return;
    state.installPrompt = null; $("nativeInstallButton").hidden = true;
    await prompt.prompt(); const choice = await prompt.userChoice;
    if (choice.outcome === "accepted") $("installDialog").close();
  });
  $("acquireCardButton").addEventListener("click", startPaidCard);
  $("nextCreateStep").addEventListener("click", () => { if (validateCreationStep(1)) setCreationStep(2); });
  $("previousCreateStep").addEventListener("click", () => setCreationStep(1));
  $("raffleForm").addEventListener("input", updateCreationPreview);
  document.querySelectorAll("[data-money]").forEach((input) => { formatMoneyField(input); input.addEventListener("input", () => { formatMoneyField(input); if (input.id === "editPremio") $("prizePreview").textContent = input.value || "Tu próximo premio"; }); });
  $("copyQrKey").addEventListener("click", () => copyText(PAYMENT_KEY, "Llave copiada: 320 875 9798"));
  $("saveQrImage").addEventListener("click", savePaymentQr);
  qrAssetReady = preloadPaymentQr();
  setupModalScrolling(); resetCreation(); renderExperience();
}
document.addEventListener("DOMContentLoaded", initExperience);
