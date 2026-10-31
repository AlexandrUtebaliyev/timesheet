// sw.js — кеширует только оболочку приложения (статику). Запросы к Apps Script
// (другой источник) не перехватываются — для них нужна сеть, офлайн-очередь в queue.js.

const CACHE_NAME = "timesheet-shell-v1";
const SHELL_FILES = [
  "./",
  "index.html",
  "styles.css",
  "config.js",
  "api.js",
  "model.js",
  "queue.js",
  "app.js",
  "demo-data.js",
  "manifest.webmanifest",
  "icons/icon-192.png",
  "icons/icon-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((names) => Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (url.origin !== location.origin || event.request.method !== "GET") return;

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request)
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          return response;
        })
        .catch(() => caches.match("index.html"));
    })
  );
});
