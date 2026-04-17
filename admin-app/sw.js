const CACHE_NAME = "mayumi-admin-survey-v72";
const ASSET_VERSION = "20260417-13";
const APP_ASSETS = [
  "./",
  "./index.html",
  `./styles.css?v=${ASSET_VERSION}`,
  `./app.js?v=${ASSET_VERSION}`,
  `./manifest.webmanifest?v=${ASSET_VERSION}`,
  `./icons/icon-192.png?v=${ASSET_VERSION}`,
  `./icons/icon-512.png?v=${ASSET_VERSION}`,
  `../shared/assets/bijiris-stamp.png?v=${ASSET_VERSION}`,
  `../default-surveys.js?v=${ASSET_VERSION}`,
  `../shared/api.js?v=${ASSET_VERSION}`,
];
const APP_ASSET_URLS = new Set(APP_ASSETS.map((asset) => new URL(asset, self.location.href).toString()));

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_ASSETS)));
  self.skipWaiting();
});

self.addEventListener("message", (event) => {
  if (event.data?.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))),
      ),
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (event.request.method !== "GET" || url.pathname.includes("/api/")) return;
  if (url.pathname.endsWith("/shared/gas-config.js")) return;

  if (event.request.mode === "navigate") {
    event.respondWith(fetch(event.request).catch(() => caches.match("./index.html")));
    return;
  }

  if (url.origin !== self.location.origin) return;
  if (!APP_ASSET_URLS.has(url.toString())) return;

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request)
        .then((response) => {
          if (!response || response.status !== 200 || response.type === "opaque") {
            return response;
          }
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          return response;
        })
        .catch(() => caches.match(event.request));
    }),
  );
});
