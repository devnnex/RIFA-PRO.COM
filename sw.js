/* Only the application shell is cached. Reservation/admin API responses never are. */
const CACHE = "rifapro-shell-v2.4.0";
const ASSETS = ["./", "./index.html", "./style.css", "./premium.css", "./app.js", "./cartones.js", "./config.js", "./manifest.json", "./images/logo.svg", "./images/icon-192.png", "./images/icon-512.png", "./images/icon-maskable-512.png", "./images/apple-touch-icon.png"];
self.addEventListener("install", (event) => { event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(ASSETS.map((asset) => new Request(new URL(asset, self.location.href), { cache: "reload" }))))); self.skipWaiting(); });
self.addEventListener("activate", (event) => { event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key.startsWith("rifapro-shell-") && key !== CACHE).map((key) => caches.delete(key)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (event.request.method !== "GET" || url.origin !== self.location.origin) return;
  const root = new URL("./", self.location.href);
  const path = url.pathname;
  if (!path.startsWith(root.pathname)) return;
  const relative = "./" + path.slice(root.pathname.length);
  if (!ASSETS.includes(relative)) return;
  // Each service-worker release installs a complete shell atomically. Reads use
  // that release immediately; no network round trip blocks an installed PWA.
  event.respondWith(caches.match(path, { cacheName: CACHE }).then((cached) => cached || fetch(event.request).then((response) => {
    if (response.ok) { const copy = response.clone(); event.waitUntil(caches.open(CACHE).then((cache) => cache.put(path, copy))); }
    return response;
  })).catch(() => event.request.mode === "navigate" ? caches.match(new URL("./index.html", self.location.href).pathname, { cacheName: CACHE }) : Response.error()));
});
