const STORAGE_RIFAS = "rifapro.rifas";
const STORAGE_QUEUE = "rifapro.syncQueue";
const STORAGE_THEME = "rifapro.theme";
const DEFAULT_SYNC_URL = "https://script.google.com/macros/s/AKfycbxYlC9MHAzYz6prdH32ZQZhZVw6U9eWkRssrJYrGHy8FPe3Zwwm1Ne9ftMrIf7OiYzc/exec";
const LIVE_SYNC_INTERVAL_MS = 2000;

const MOBILE_BREAKPOINT = 768;
const MOBILE_PAGE_SIZE = 40;

let rifas = [];
let rifaActivaId = null;
let numeroActualIndex = null;
let pagina = 1;
let flushingQueue = false;
let loaderCounter = 0;
let syncInFlight = false;

const syncState = {
  url: "",
  queue: [],
  lastStatus: ""
};
let pinModalResolver = null;
let pinModalCleanup = null;

const els = {};

function init() {
  mapElements();
  bindEvents();

  rifas = readJSON(STORAGE_RIFAS, []);
  syncState.url = DEFAULT_SYNC_URL;
  syncState.queue = readJSON(STORAGE_QUEUE, []);
  applyTheme(localStorage.getItem(STORAGE_THEME) || "light");

  renderLista();
  renderRifa();

  window.addEventListener("online", () => {
    setSyncStatus("Conexion restablecida, sincronizando...");
    sincronizarTodo();
  });

  window.addEventListener("resize", () => {
    if (window.innerWidth > MOBILE_BREAKPOINT) {
      closeSidebar();
    }
    if (rifaActivaId) {
      renderRifa();
    }
  });

  if (syncState.url) {
    sincronizarTodo(true);
    window.setInterval(() => {
      sincronizarTodo(false);
    }, LIVE_SYNC_INTERVAL_MS);
  }
}

function mapElements() {
  const ids = [
    "menuBtn", "sidebar", "sidebarBackdrop", "btnCerrarSidebar", "listaRifas", "titulo", "precio", "cantidad", "premio", "loteria", "fecha", "hora", "responsable",
    "btnCrearRifa", "themeToggle", "themeIcon", "themeLabel", "infoRifa", "kpiLibres", "kpiApartados",
    "kpiPagados", "kpiRecaudado", "barra", "porcentajeTexto", "dashboard", "paginacion", "modal", "clienteNombre",
    "clienteTelefono", "estadoSelect", "btnGuardarNumero", "btnLiberarNumero", "btnCerrarModal",
    "pinModal", "pinTitle", "pinHelp", "pinInput", "btnPinConfirmar", "btnPinCancelar", "pinBtnSpinner", "pinBtnText",
    "appLoader", "appLoaderText"
  ];

  ids.forEach((id) => {
    els[id] = document.getElementById(id);
  });
}

function bindEvents() {
  els.menuBtn.addEventListener("click", toggleMenu);
  els.btnCerrarSidebar.addEventListener("click", closeSidebar);
  els.sidebarBackdrop.addEventListener("click", closeSidebar);
  els.btnCrearRifa.addEventListener("click", crearRifa);
  els.themeToggle.addEventListener("click", toggleTheme);
  els.btnGuardarNumero.addEventListener("click", guardarCliente);
  els.btnLiberarNumero.addEventListener("click", liberarNumero);
  els.btnCerrarModal.addEventListener("click", cerrarModal);
  els.precio.addEventListener("input", () => {
    els.precio.value = formatMoneyInput(els.precio.value);
  });
  els.precio.addEventListener("blur", () => {
    els.precio.value = formatMoneyInput(els.precio.value);
  });

  els.modal.addEventListener("click", (event) => {
    if (event.target === els.modal) {
      cerrarModal();
    }
  });

  els.pinModal.addEventListener("click", (event) => {
    if (event.target === els.pinModal) {
      closePinModal(null, true);
    }
  });
}

function readJSON(key, fallback) {
  try {
    return JSON.parse(localStorage.getItem(key)) || fallback;
  } catch (error) {
    return fallback;
  }
}

function writeRifas() {
  localStorage.setItem(STORAGE_RIFAS, JSON.stringify(rifas));
}

function writeQueue() {
  localStorage.setItem(STORAGE_QUEUE, JSON.stringify(syncState.queue));
}

function showAppLoader(message = "Cargando datos...") {
  loaderCounter += 1;
  els.appLoaderText.textContent = message;
  els.appLoader.classList.add("show");
  els.appLoader.setAttribute("aria-hidden", "false");
}

function hideAppLoader() {
  loaderCounter = Math.max(0, loaderCounter - 1);
  if (loaderCounter > 0) return;
  els.appLoader.classList.remove("show");
  els.appLoader.setAttribute("aria-hidden", "true");
}

function normalizeText(value) {
  return (value || "").toString().trim();
}

function normalizeNumeroId(value) {
  const raw = normalizeText(value);
  if (!raw) return "";
  if (/^\d+$/.test(raw)) return String(Number(raw));
  return raw;
}

function escapeHtml(value) {
  return (value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function formatoCOP(valor) {
  return new Intl.NumberFormat("es-CO", {
    style: "currency",
    currency: "COP",
    maximumFractionDigits: 0
  }).format(valor || 0);
}

function toggleMenu() {
  if (els.sidebar.classList.contains("open")) {
    closeSidebar();
    return;
  }
  openSidebar();
}

function openSidebar() {
  els.sidebar.classList.add("open");
  els.sidebarBackdrop.classList.add("show");
  els.sidebarBackdrop.setAttribute("aria-hidden", "false");
}

function closeSidebar() {
  els.sidebar.classList.remove("open");
  els.sidebarBackdrop.classList.remove("show");
  els.sidebarBackdrop.setAttribute("aria-hidden", "true");
}

function parseMoneyValue(raw) {
  const digits = String(raw || "").replace(/\D/g, "");
  return Number(digits || 0);
}

function formatMoneyInput(raw) {
  const value = parseMoneyValue(raw);
  if (!value) return "";
  return `$${value.toLocaleString("es-CO")}`;
}

function formatFechaBonita(rawDate) {
  const safe = normalizeText(rawDate);
  if (!safe) return "";

  let date;
  const ymd = safe.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ymd) {
    // Parse YYYY-MM-DD in local time to avoid timezone shifts.
    date = new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
  } else {
    // Accept ISO and other valid date strings coming from Sheets.
    date = new Date(safe);
  }

  if (Number.isNaN(date.getTime())) return safe;

  const formatted = new Intl.DateTimeFormat("es-CO", {
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric"
  }).format(date);

  return formatted.charAt(0).toUpperCase() + formatted.slice(1);
}

function formatHoraBonita(rawTime) {
  const safe = normalizeText(rawTime);
  if (!safe) return "";

  let date;
  const hhmm = safe.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (hhmm) {
    const hours = Number(hhmm[1]);
    const minutes = Number(hhmm[2]);
    if (!Number.isInteger(hours) || !Number.isInteger(minutes)) return safe;
    date = new Date();
    date.setHours(hours, minutes, 0, 0);
  } else {
    const parsed = new Date(safe);
    if (Number.isNaN(parsed.getTime())) return safe;
    date = parsed;
  }

  return new Intl.DateTimeFormat("es-CO", {
    hour: "numeric",
    minute: "2-digit",
    hour12: true
  }).format(date);
}

function applyTheme(mode) {
  const darkMode = mode === "dark";
  document.body.classList.toggle("dark-mode", darkMode);
  localStorage.setItem(STORAGE_THEME, darkMode ? "dark" : "light");
  els.themeLabel.textContent = darkMode ? "Modo dia" : "Modo noche";
  els.themeIcon.className = darkMode ? "fa-solid fa-sun" : "fa-solid fa-moon";
}

function toggleTheme() {
  const next = document.body.classList.contains("dark-mode") ? "light" : "dark";
  applyTheme(next);
}

function createNumeros(cantidad, rifaId, defaultUpdatedAt = "") {
  const safeCantidad = Math.max(1, Number(cantidad) || 1);
  const padding = Math.max(2, String(safeCantidad - 1).length);
  const numbers = [];

  for (let i = 0; i < safeCantidad; i += 1) {
    const numeroId = String(i).padStart(padding, "0");
    numbers.push({
      id: `${rifaId}-${numeroId}`,
      n: numeroId,
      estado: "libre",
      cliente: null,
      updatedAt: defaultUpdatedAt
    });
  }

  return numbers;
}

function crearRifaDesdeMeta(meta) {
  const rifaId = String(meta.rifaId);
  const cantidad = Math.max(1, Number(meta.cantidad) || 100);

  return {
    id: rifaId,
    titulo: normalizeText(meta.titulo) || `Rifa ${rifaId}`,
    precio: Number(meta.precio) || 0,
    cantidad,
    premio: normalizeText(meta.premio),
    loteria: normalizeText(meta.loteria),
    fecha: normalizeText(meta.fecha),
    hora: normalizeText(meta.hora),
    responsable: normalizeText(meta.responsable),
    createdAt: new Date().toISOString(),
    numeros: createNumeros(cantidad, rifaId, "")
  };
}

async function crearRifa() {
  const autorizado = await validarPinAdmin("Ingresa PIN de 4 digitos para crear una rifa.");
  if (!autorizado) return;

  const cantidad = Number(els.cantidad.value);
  const precio = parseMoneyValue(els.precio.value);
  const nueva = {
    id: Date.now().toString(),
    titulo: normalizeText(els.titulo.value),
    precio,
    cantidad,
    premio: normalizeText(els.premio.value),
    loteria: normalizeText(els.loteria.value),
    fecha: normalizeText(els.fecha.value),
    hora: normalizeText(els.hora.value),
    responsable: normalizeText(els.responsable.value),
    createdAt: new Date().toISOString(),
    numeros: []
  };

  if (!nueva.titulo || !nueva.precio || !nueva.cantidad) {
    alert("Completa titulo, precio y cantidad.");
    return;
  }

  if (nueva.cantidad > 100000) {
    alert("La cantidad maxima soportada es 100000.");
    return;
  }

  nueva.numeros = createNumeros(nueva.cantidad, nueva.id, new Date().toISOString());

  rifas.unshift(nueva);
  rifaActivaId = nueva.id;
  pagina = 1;

  writeRifas();
  renderLista();
  renderRifa();
  limpiarFormularioRifa();

  enqueueOperation({
    type: "upsertRifaMeta",
    payload: buildRifaMetaPayload(nueva)
  });
  flushQueue();
}

function limpiarFormularioRifa() {
  ["titulo", "precio", "cantidad", "premio", "loteria", "fecha", "hora", "responsable"].forEach((id) => {
    els[id].value = "";
  });
}

function getRifaActiva() {
  return rifas.find((rifa) => rifa.id === rifaActivaId) || null;
}

function renderLista() {
  els.listaRifas.innerHTML = "";

  rifas.forEach((rifa) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = `rifa-item ${rifaActivaId === rifa.id ? "active" : ""}`;

    const title = document.createElement("span");
    title.className = "rifa-title";
    title.textContent = rifa.titulo;

    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "delete-btn";
    delBtn.innerHTML = '<i class="fa-solid fa-trash"></i>';
    delBtn.addEventListener("click", (event) => {
      event.stopPropagation();
      eliminarRifa(rifa.id);
    });

    item.addEventListener("click", () => abrirRifa(rifa.id));

    item.appendChild(title);
    item.appendChild(delBtn);
    els.listaRifas.appendChild(item);
  });
}

async function eliminarRifa(id) {
  const autorizado = await validarPinAdmin("Ingresa PIN de 4 digitos para eliminar esta rifa.");
  if (!autorizado) return;

  rifas = rifas.filter((rifa) => rifa.id !== id);
  if (rifaActivaId === id) {
    rifaActivaId = null;
  }

  writeRifas();
  renderLista();
  renderRifa();

  enqueueOperation({
    type: "deleteRifa",
    payload: { rifaId: id }
  });
  flushQueue();
}

async function abrirRifa(id) {
  rifaActivaId = id;
  pagina = 1;
  renderLista();
  renderRifa();

  const rifa = getRifaActiva();
  if (rifa) {
    showAppLoader("Cargando casillas ocupadas...");
    try {
      await cargarRifaDesdeSheets(rifa);
    } finally {
      hideAppLoader();
    }
  }

  if (window.innerWidth <= MOBILE_BREAKPOINT) {
    closeSidebar();
  }
}

function renderRifa() {
  const rifa = getRifaActiva();

  if (!rifa) {
    els.infoRifa.innerHTML = "<h2>Selecciona o crea una rifa</h2>";
    els.dashboard.innerHTML = "";
    els.paginacion.innerHTML = "";
    ["kpiLibres", "kpiApartados", "kpiPagados", "kpiRecaudado"].forEach((id) => {
      els[id].textContent = "";
    });
    els.barra.style.width = "0%";
    els.porcentajeTexto.textContent = "";
    return;
  }

  els.infoRifa.innerHTML = `
    <h2>${escapeHtml(rifa.titulo)}</h2>
    <p><strong>Premio:</strong> ${escapeHtml(rifa.premio || "-")}</p>
    <p><strong>Loteria:</strong> ${escapeHtml(rifa.loteria || "-")}</p>
    <p><strong>Fecha:</strong> ${escapeHtml(formatFechaBonita(rifa.fecha) || "-")} ${escapeHtml(formatHoraBonita(rifa.hora) || "")}</p>
    <p><strong>Responsable:</strong> ${escapeHtml(rifa.responsable || "-")}</p>
  `;

  renderDashboard(rifa);
  actualizarKPIs(rifa);
}

function renderDashboard(rifa) {
  const isMobile = window.innerWidth <= MOBILE_BREAKPOINT;
  const total = rifa.numeros.length;

  let start = 0;
  let end = total;

  if (isMobile) {
    start = (pagina - 1) * MOBILE_PAGE_SIZE;
    end = Math.min(start + MOBILE_PAGE_SIZE, total);
    renderPaginacion(total);
  } else {
    els.paginacion.innerHTML = "";
  }

  els.dashboard.innerHTML = "";

  for (let index = start; index < end; index += 1) {
    const numero = rifa.numeros[index];
    const tile = document.createElement("div");
    tile.className = `numero ${numero.estado}`;

    const nombre = numero.cliente?.nombre || "";
    const nombreCorto = nombre.length > 10 ? `${nombre.slice(0, 10)}...` : nombre;

    tile.innerHTML = `
      <span class="num">${escapeHtml(numero.n)}</span>
      ${nombre ? `<span class="cliente">${escapeHtml(nombreCorto)}</span>` : ""}
    `;

    tile.title = nombre;
    tile.addEventListener("click", () => abrirModal(index));
    els.dashboard.appendChild(tile);
  }
}

function renderPaginacion(totalNumeros) {
  const totalPaginas = Math.ceil(totalNumeros / MOBILE_PAGE_SIZE);

  els.paginacion.innerHTML = "";

  if (totalPaginas <= 1) {
    return;
  }

  const prevBtn = document.createElement("button");
  prevBtn.type = "button";
  prevBtn.className = "pager-arrow";
  prevBtn.innerHTML = "&#8249;";
  prevBtn.disabled = pagina <= 1;
  prevBtn.addEventListener("click", () => {
    if (pagina > 1) {
      pagina -= 1;
      renderRifa();
    }
  });

  const dots = document.createElement("span");
  dots.className = "pager-dots";
  const activeDot = totalPaginas <= 1
    ? 1
    : Math.min(2, Math.max(0, Math.round(((pagina - 1) / (totalPaginas - 1)) * 2)));

  for (let i = 0; i < 3; i += 1) {
    const dot = document.createElement("span");
    dot.className = `pager-dot ${i === activeDot ? "active" : ""}`;
    dots.appendChild(dot);
  }

  const nextBtn = document.createElement("button");
  nextBtn.type = "button";
  nextBtn.className = "pager-arrow";
  nextBtn.innerHTML = "&#8250;";
  nextBtn.disabled = pagina >= totalPaginas;
  nextBtn.addEventListener("click", () => {
    if (pagina < totalPaginas) {
      pagina += 1;
      renderRifa();
    }
  });

  els.paginacion.appendChild(prevBtn);
  els.paginacion.appendChild(dots);
  els.paginacion.appendChild(nextBtn);
}

function actualizarKPIs(rifa) {
  const libres = rifa.numeros.filter((n) => n.estado === "libre").length;
  const apartados = rifa.numeros.filter((n) => n.estado === "apartado").length;
  const pagados = rifa.numeros.filter((n) => n.estado === "pagado").length;

  const recaudado = pagados * rifa.precio;
  const meta = rifa.cantidad * rifa.precio;
  const avance = meta ? (recaudado / meta) * 100 : 0;
  const porcentajeSeguro = Math.max(0, Math.min(100, avance));

  els.kpiLibres.innerHTML = `<strong>Libres</strong>${libres}`;
  els.kpiApartados.innerHTML = `<strong>Apartados</strong>${apartados}`;
  els.kpiPagados.innerHTML = `<strong>Pagados</strong>${pagados}`;
  els.kpiRecaudado.innerHTML = `<strong>Recaudado</strong>${formatoCOP(recaudado)}`;

  els.barra.style.width = `${porcentajeSeguro.toFixed(1)}%`;
  els.porcentajeTexto.textContent = `${porcentajeSeguro.toFixed(1)}% completado`;
}

async function abrirModal(index) {
  const rifa = getRifaActiva();
  if (!rifa) return;

  numeroActualIndex = index;
  const num = rifa.numeros[index];

  if (num.estado !== "libre") {
    const autorizado = await validarPinEdicion();
    if (!autorizado) {
      numeroActualIndex = null;
      return;
    }
  }

  els.clienteNombre.value = num.cliente?.nombre || "";
  els.clienteTelefono.value = num.cliente?.telefono || "";
  els.estadoSelect.value = num.estado;

  els.modal.classList.add("show");
  els.modal.setAttribute("aria-hidden", "false");
}

async function validarPinEdicion() {
  if (!syncState.url || !navigator.onLine) {
    alert("Se requiere conexion para validar PIN.");
    return false;
  }

  const pinRaw = await openPinModal({
    title: "Validacion de seguridad",
    help: "Este numero ya esta apartado o pagado. Ingresa PIN de 4 digitos."
  });
  if (pinRaw === null) {
    return false;
  }

  const pin = String(pinRaw).replace(/\D/g, "").slice(0, 4);
  if (pin.length !== 4) {
    closePinModal();
    alert("PIN invalido. Debe tener 4 digitos.");
    return false;
  }

  return verificarPinYFinalizar(pin);
}

async function validarPinAdmin(helpText) {
  if (!syncState.url || !navigator.onLine) {
    alert("Se requiere conexion para validar PIN.");
    return false;
  }

  const pinRaw = await openPinModal({
    title: "Validacion de seguridad",
    help: helpText
  });
  if (pinRaw === null) {
    return false;
  }

  const pin = String(pinRaw).replace(/\D/g, "").slice(0, 4);
  if (pin.length !== 4) {
    closePinModal();
    alert("PIN invalido. Debe tener 4 digitos.");
    return false;
  }

  return verificarPinYFinalizar(pin);
}

async function verificarPinYFinalizar(pin) {
  try {
    setPinLoading(true);
    const response = await apiGet("verifyPin", { payload: JSON.stringify({ pin }) });
    setPinLoading(false);
    if (response && response.ok && response.valid === true) {
      closePinModal();
      return true;
    }

    closePinModal();
    alert("PIN incorrecto.");
    return false;
  } catch (error) {
    setPinLoading(false);
    closePinModal();
    alert("No fue posible validar el PIN.");
    return false;
  }
}

function openPinModal(options = {}) {
  return new Promise((resolve) => {
    pinModalResolver = resolve;
    if (pinModalCleanup) {
      pinModalCleanup();
      pinModalCleanup = null;
    }
    els.pinInput.value = "";
    setPinLoading(false);
    els.pinTitle.textContent = options.title || "Validacion de seguridad";
    els.pinHelp.textContent = options.help || "Ingresa PIN de 4 digitos.";
    els.pinModal.classList.add("show");
    els.pinModal.setAttribute("aria-hidden", "false");
    els.pinInput.focus();

    const onConfirm = () => {
      const value = els.pinInput.value || "";
      cleanup();
      if (pinModalResolver) {
        pinModalResolver(value);
        pinModalResolver = null;
      }
    };

    const onCancel = () => {
      closePinModal(null, true);
      cleanup();
    };

    const onEnter = (event) => {
      if (event.key === "Enter") {
        onConfirm();
      }
      if (event.key === "Escape") {
        onCancel();
      }
    };

    const cleanup = () => {
      els.btnPinConfirmar.removeEventListener("click", onConfirm);
      els.btnPinCancelar.removeEventListener("click", onCancel);
      els.pinInput.removeEventListener("keydown", onEnter);
      pinModalCleanup = null;
    };
    pinModalCleanup = cleanup;

    els.btnPinConfirmar.addEventListener("click", onConfirm);
    els.btnPinCancelar.addEventListener("click", onCancel);
    els.pinInput.addEventListener("keydown", onEnter);
  });
}

function closePinModal(result = null, shouldResolve = false) {
  els.pinModal.classList.remove("show");
  els.pinModal.setAttribute("aria-hidden", "true");
  setPinLoading(false);

  if (pinModalCleanup) {
    pinModalCleanup();
  }

  if (shouldResolve && pinModalResolver) {
    pinModalResolver(result);
    pinModalResolver = null;
  }
}

function setPinLoading(isLoading) {
  els.btnPinConfirmar.disabled = isLoading;
  els.btnPinCancelar.disabled = isLoading;
  els.pinInput.disabled = isLoading;
  els.pinBtnSpinner.classList.toggle("show", isLoading);
  els.pinBtnText.textContent = isLoading ? "Validando..." : "Validar";
}

function cerrarModal() {
  els.modal.classList.remove("show");
  els.modal.setAttribute("aria-hidden", "true");
  numeroActualIndex = null;
}

function guardarCliente() {
  const rifa = getRifaActiva();
  if (!rifa || numeroActualIndex === null) return;

  const estado = els.estadoSelect.value;
  const nombre = normalizeText(els.clienteNombre.value);
  const telefono = normalizeText(els.clienteTelefono.value);

  const num = rifa.numeros[numeroActualIndex];
  num.estado = estado;
  num.updatedAt = new Date().toISOString();
  num.cliente = estado === "libre" ? null : { nombre, telefono };

  persistirRifaActiva();
  cerrarModal();

  upsertNumeroRemoto(rifa, num);
}

function liberarNumero() {
  const rifa = getRifaActiva();
  if (!rifa || numeroActualIndex === null) return;

  const num = rifa.numeros[numeroActualIndex];
  num.estado = "libre";
  num.cliente = null;
  num.updatedAt = new Date().toISOString();

  persistirRifaActiva();
  cerrarModal();

  upsertNumeroRemoto(rifa, num);
}

function persistirRifaActiva() {
  const rifa = getRifaActiva();
  if (!rifa) return;

  rifas = rifas.map((item) => (item.id === rifa.id ? rifa : item));
  writeRifas();
  renderRifa();
  renderLista();
}

function setSyncStatus(message) {
  syncState.lastStatus = message;
  console.log("[RifaPro Sync]", message);
}

function buildRifaMetaPayload(rifa) {
  return {
    rifaId: rifa.id,
    titulo: rifa.titulo,
    precio: rifa.precio,
    premio: rifa.premio,
    loteria: rifa.loteria,
    fecha: rifa.fecha,
    hora: rifa.hora,
    responsable: rifa.responsable,
    cantidad: rifa.cantidad,
    updatedAt: new Date().toISOString()
  };
}

function buildNumeroPayload(rifa, numero) {
  return {
    rifaId: rifa.id,
    numeroId: numero.n,
    estado: numero.estado,
    clienteNombre: numero.cliente?.nombre || "",
    clienteTelefono: numero.cliente?.telefono || "",
    updatedAt: numero.updatedAt || new Date().toISOString(),
    titulo: rifa.titulo,
    precio: rifa.precio,
    premio: rifa.premio,
    loteria: rifa.loteria,
    fecha: rifa.fecha,
    hora: rifa.hora,
    responsable: rifa.responsable,
    cantidad: rifa.cantidad
  };
}

function enqueueOperation(op) {
  const opWithMeta = {
    ...op,
    attempts: 0,
    lastError: "",
    queuedAt: new Date().toISOString()
  };

  if (op.type === "upsertNumber") {
    const key = `${op.payload.rifaId}-${op.payload.numeroId}`;
    syncState.queue = syncState.queue.filter((item) => {
      if (item.type !== "upsertNumber") return true;
      const existingKey = `${item.payload.rifaId}-${item.payload.numeroId}`;
      return existingKey !== key;
    });
  }

  syncState.queue.push(opWithMeta);
  writeQueue();
}

async function flushQueue() {
  if (!syncState.url || flushingQueue || !navigator.onLine) {
    return;
  }

  flushingQueue = true;

  try {
    const pending = [...syncState.queue];
    syncState.queue = [];

    for (const op of pending) {
      try {
        if (op.type === "upsertNumber") {
          await apiPost("upsertNumber", op.payload);
        }

        if (op.type === "upsertRifaMeta") {
          await apiPost("upsertRifaMeta", op.payload);
        }

        if (op.type === "deleteRifa") {
          await apiPost("deleteRifa", op.payload);
        }
      } catch (error) {
        op.attempts = Number(op.attempts || 0) + 1;
        op.lastError = String(error.message || error);
        syncState.queue.push(op);
      }
    }

    writeQueue();
    if (syncState.queue.length) {
      setSyncStatus(`Sync parcial. Pendientes: ${syncState.queue.length}`);
    } else {
      setSyncStatus("Sync al dia");
    }
  } finally {
    flushingQueue = false;
  }
}

async function upsertNumeroRemoto(rifa, numero) {
  const payload = buildNumeroPayload(rifa, numero);

  if (!syncState.url || !navigator.onLine) {
    enqueueOperation({ type: "upsertNumber", payload });
    setSyncStatus("Cambio en cola local. Sincroniza al reconectar.");
    return;
  }

  try {
    await apiPost("upsertNumber", payload);
    setSyncStatus(`Numero ${numero.n} sincronizado`);
    flushQueue();
  } catch (error) {
    enqueueOperation({ type: "upsertNumber", payload });
    setSyncStatus("Error en sync, cambio guardado en cola local.");
  }
}

async function cargarRifasDesdeSheets() {
  if (!syncState.url || !navigator.onLine) {
    return;
  }

  const response = await apiGet("getRifas");
  if (!response.ok || !Array.isArray(response.data)) {
    return;
  }

  let changed = false;
  const localMap = new Map(rifas.map((item) => [item.id, item]));

  response.data.forEach((meta) => {
    const id = String(meta.rifaId || "");
    if (!id) return;

    const local = localMap.get(id);
    if (!local) {
      const nueva = crearRifaDesdeMeta(meta);
      rifas.unshift(nueva);
      localMap.set(id, nueva);
      changed = true;
      return;
    }

    const remoteCantidad = Math.max(1, Number(meta.cantidad) || local.cantidad || 1);
    if (Number(local.cantidad) !== remoteCantidad) {
      local.cantidad = remoteCantidad;
      if (!Array.isArray(local.numeros) || local.numeros.length !== remoteCantidad) {
        local.numeros = createNumeros(remoteCantidad, local.id);
      }
      changed = true;
    }

    ["titulo", "premio", "loteria", "fecha", "hora", "responsable"].forEach((field) => {
      const remoteValue = normalizeText(meta[field]);
      if (remoteValue && local[field] !== remoteValue) {
        local[field] = remoteValue;
        changed = true;
      }
    });

    const remotePrecio = Number(meta.precio);
    if (Number.isFinite(remotePrecio) && remotePrecio >= 0 && local.precio !== remotePrecio) {
      local.precio = remotePrecio;
      changed = true;
    }
  });

  if (changed) {
    writeRifas();
    renderLista();

    if (!rifaActivaId && rifas.length) {
      rifaActivaId = rifas[0].id;
      pagina = 1;
    }

    renderRifa();
  }
}

async function cargarRifaDesdeSheets(rifa) {
  if (!syncState.url || !navigator.onLine) {
    return;
  }

  try {
    const response = await apiGet("getRifaNumbers", { rifaId: rifa.id });
    if (!response.ok || !Array.isArray(response.data)) {
      return;
    }

    const mapaLocal = new Map(rifa.numeros.map((n) => [normalizeNumeroId(n.n), n]));
    let cambios = 0;

    response.data.forEach((record) => {
      const numeroId = normalizeNumeroId(record.numeroId);
      if (!numeroId) return;

      const local = mapaLocal.get(numeroId);
      if (!local) return;

      const remoteDate = record.updatedAt ? new Date(record.updatedAt).getTime() : 0;
      const localDate = local.updatedAt ? new Date(local.updatedAt).getTime() : 0;
      const localEsInicial = local.estado === "libre" && !local.cliente;

      if (localEsInicial || remoteDate >= localDate) {
        local.estado = record.estado || "libre";
        local.updatedAt = record.updatedAt || new Date().toISOString();
        local.cliente = local.estado === "libre"
          ? null
          : {
            nombre: record.clienteNombre || "",
            telefono: record.clienteTelefono || ""
          };
        cambios += 1;
      }
    });

    if (cambios > 0) {
      persistirRifaActiva();
      setSyncStatus(`Rifa actualizada desde Sheets (${cambios} cambios)`);
    }
  } catch (error) {
    setSyncStatus("No se pudo leer Google Sheets.");
  }
}

async function sincronizarTodo(showLoader = false) {
  if (!syncState.url) {
    setSyncStatus("Configura primero la URL de Apps Script.");
    return;
  }
  if (syncInFlight) {
    return;
  }
  syncInFlight = true;

  if (showLoader) {
    showAppLoader("Cargando casillas ocupadas...");
  }

  setSyncStatus("Sincronizando...");

  try {
    await cargarRifasDesdeSheets();

    if (!rifaActivaId && rifas.length) {
      rifaActivaId = rifas[0].id;
      pagina = 1;
    }

    const rifa = getRifaActiva();
    if (rifa) {
      await cargarRifaDesdeSheets(rifa);
    }

    renderLista();
    renderRifa();

    await flushQueue();
    if (!syncState.queue.length) {
      setSyncStatus("Sync al dia");
    }
  } catch (error) {
    setSyncStatus("Error de conexion con Apps Script.");
  } finally {
    syncInFlight = false;
    if (showLoader) {
      hideAppLoader();
    }
  }
}

async function apiGet(action, params = {}) {
  return jsonpRequest(action, params);
}

async function apiPost(action, payload = {}) {
  const packedPayload = JSON.stringify(payload);

  try {
    const result = await jsonpRequest(action, { payload: packedPayload });
    if (result && result.ok === false) {
      throw new Error(result.error || "jsonp action rejected");
    }
    return result || { ok: true, transport: "jsonp" };
  } catch (error) {
    return postFallbackNoCors(action, packedPayload);
  }
}

async function postFallbackNoCors(action, packedPayload) {
  const form = new URLSearchParams();
  form.set("action", action);
  form.set("payload", packedPayload);

  await fetch(syncState.url, {
    method: "POST",
    mode: "no-cors",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"
    },
    body: form.toString()
  });

  return { ok: true, transport: "post-no-cors" };
}

function jsonpRequest(action, params = {}) {
  return new Promise((resolve, reject) => {
    if (!syncState.url) {
      reject(new Error("sync URL vacia"));
      return;
    }

    const callbackId = `rifaJsonp_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
    const url = new URL(syncState.url);
    url.searchParams.set("action", action);
    url.searchParams.set("callback", callbackId);

    Object.entries(params).forEach(([key, value]) => {
      if (value !== undefined && value !== null) {
        url.searchParams.set(key, String(value));
      }
    });

    const script = document.createElement("script");
    script.src = url.toString();
    script.async = true;

    const cleanup = () => {
      if (window[callbackId]) {
        delete window[callbackId];
      }
      if (script.parentNode) {
        script.parentNode.removeChild(script);
      }
    };

    const timer = setTimeout(() => {
      cleanup();
      reject(new Error("timeout jsonp"));
    }, 12000);

    window[callbackId] = (data) => {
      clearTimeout(timer);
      cleanup();
      resolve(data);
    };

    script.onerror = () => {
      clearTimeout(timer);
      cleanup();
      reject(new Error("jsonp network error"));
    };

    document.body.appendChild(script);
  });
}

document.addEventListener("DOMContentLoaded", init);
