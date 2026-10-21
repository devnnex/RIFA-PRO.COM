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
