// MB Notebook offline support: keeps the app itself available without internet.
// Your pages are never cached here (they're kept in the app's own offline storage).
const CACHE = "mbn-2.0.2";
const SHELL = ["./", "./index.html", "./app.js", "./config.js", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png"];
self.addEventListener("install", (e) => { e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.hostname.endsWith("googleapis.com") && url.hostname !== "fonts.googleapis.com") return;   // Drive API: always live
  if (url.hostname === "accounts.google.com" && !url.pathname.startsWith("/gsi/client")) return;
  const fromCdn = url.hostname === "cdn.jsdelivr.net" || url.hostname === "fonts.gstatic.com";
  if (fromCdn) {   // versioned files: cache first
    e.respondWith(caches.match(req).then((hit) => hit || fetch(req).then((res) => { if (res.ok || res.type === "opaque") { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); } return res; })));
    return;
  }
  // app files, Google fonts CSS, Google sign-in script: network first so updates arrive, cache when offline
  e.respondWith(fetch(req).then((res) => { if (res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); } return res; })
    .catch(() => caches.match(req).then((hit) => hit || (req.mode === "navigate" ? caches.match("./index.html") : Response.error()))));
});
