const ADMIN_CACHE_NAME = 'mayumi-admin-shell-v2';
const ADMIN_ENTRY_CANDIDATES = [
  './mayumi-admin.html',
  './index.html'
];
const ADMIN_SHELL_ASSETS = [
  './admin-manifest.json',
  './admin-icon-192.png',
  './admin-icon-512.png',
  './admin-apple-touch-icon.png',
  './assets/icon.png',
  './assets/instr_ios_1.png',
  './assets/instr_ios_2.png',
  './assets/instr_ios_3.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(ADMIN_CACHE_NAME)
      .then(function (cache) {
        return Promise.all(
          ADMIN_SHELL_ASSETS.map(function (asset) {
            return cache.add(asset).catch(function () {
              return null;
            });
          }).concat(
            ADMIN_ENTRY_CANDIDATES.map(function (asset) {
              return fetch(asset, { cache: 'no-cache' })
                .then(function (response) {
                  if (!response || !response.ok) return null;
                  return cache.put(asset, response.clone());
                })
                .catch(function () {
                  return null;
                });
            })
          )
        );
      })
      .then(function () {
        return self.skipWaiting();
      })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (key) {
        if (key !== ADMIN_CACHE_NAME) {
          return caches.delete(key);
        }
        return null;
      }));
    }).then(function () {
      return self.clients.claim();
    })
  );
});

self.addEventListener('fetch', function (event) {
  if (!event.request || event.request.method !== 'GET') return;

  const requestUrl = new URL(event.request.url);
  if (requestUrl.origin !== self.location.origin) return;

  if (event.request.mode === 'navigate') {
    const fallbackKey = requestUrl.pathname.endsWith('/mayumi-admin.html') ? './mayumi-admin.html' : './index.html';
    event.respondWith(
      fetch(event.request)
        .then(function (response) {
          const cloned = response.clone();
          caches.open(ADMIN_CACHE_NAME).then(function (cache) {
            cache.put(fallbackKey, cloned);
          });
          return response;
        })
        .catch(function () {
          return caches.match(fallbackKey)
            .then(function (cached) {
              return cached || caches.match('./mayumi-admin.html') || caches.match('./index.html');
            });
        })
    );
    return;
  }

  if (ADMIN_SHELL_ASSETS.some(function (asset) { return requestUrl.pathname.endsWith(asset.replace('./', '/')); })) {
    event.respondWith(
      caches.match(event.request).then(function (cached) {
        if (cached) return cached;
        return fetch(event.request).then(function (response) {
          const cloned = response.clone();
          caches.open(ADMIN_CACHE_NAME).then(function (cache) {
            cache.put(event.request, cloned);
          });
          return response;
        });
      })
    );
  }
});
