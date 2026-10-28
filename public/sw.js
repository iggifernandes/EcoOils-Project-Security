// Basic service worker for the Security Timesheet app.
// Bump CACHE_NAME whenever you deploy a new version so old caches get cleared.
const CACHE_NAME = "ecooils-security-v1.2.4";

// Files that exist from day one — safe to precache. Add your built JS/CSS
// filenames here too once you know them (e.g. "/assets/index-ab12cd34.js").
const APP_SHELL = [
  ".",
  "index.html",
  "manifest.json",
  "icon-192.png",
  "icon-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

// Network-first for navigation/API-ish requests, falling back to cache when offline.
// Same-origin GETs get opportunistically cached as they're used, so assets you
// didn't list in APP_SHELL still end up available offline after first visit.
self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return; // don't touch cross-origin (e.g. storage API) calls

  event.respondWith(
    fetch(request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy)).catch(() => {});
        return response;
      })
      .catch(() => caches.match(request).then((cached) => cached || caches.match("index.html")))
  );
});
