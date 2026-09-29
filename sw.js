// Offline shell for Receipt Capture. Network first so updates on GitHub arrive straight away; cache used when offline.
const CACHE = "receipt-capture-v1";
const SHELL = ["./", "index.html", "app.js", "config.js", "redirect.html", "manifest.webmanifest",
  "lib/msal-browser.min.js", "lib/msal-redirect-bridge.min.js", "lib/piexif.js",
  "icons/icon-192.png", "icons/icon-512.png", "icons/icon-512-maskable.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== self.location.origin) return;   // never touch Microsoft / GPS lookups
  if (u.pathname.endsWith("redirect.html") && u.search) return;                  // sign-in responses always go to network
  e.respondWith(
    fetch(e.request).then((r) => {
      if (r.ok) { const copy = r.clone(); caches.open(CACHE).then((c) => c.put(e.request, copy)); }
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match("index.html")))
  );
});
