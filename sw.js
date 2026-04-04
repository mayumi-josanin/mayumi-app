const CACHE_NAME = 'mayumi-app-v6';
const ASSETS = [
  './',
  './index.html',
  './app.js',
  './style.css',
  './icon.png',
  './manifest.json'
];

function isHtmlRequest(request) {
  const accept = request.headers.get('accept') || '';
  return request.mode === 'navigate' || accept.includes('text/html');
}

async function networkFirst(request) {
  const cache = await caches.open(CACHE_NAME);
  try {
    const response = await fetch(request);
    cache.put(request, response.clone());
    return response;
  } catch (err) {
    const cached = await cache.match(request);
    if (cached) return cached;
    throw err;
  }
}

async function staleWhileRevalidate(request) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(request);
  const networkPromise = fetch(request).then((response) => {
    cache.put(request, response.clone());
    return response;
  }).catch(() => null);
  return cached || networkPromise || fetch(request);
}

self.addEventListener('push', (event) => {
  const data = event.data ? event.data.json() : { title: 'まゆみ助産院', body: '新しいお知らせがあります' };
  const targetPage = data && data.data && data.data.openPage ? String(data.data.openPage) : String(data.openPage || '');
  const targetUrl = data && data.url
    ? data.url
    : ('./index.html' + (targetPage ? ('?open=' + encodeURIComponent(targetPage)) : ''));
  const options = {
    body: data.body,
    icon: './icon.png',
    badge: './icon.png',
    vibrate: [200, 100, 200],
    data: {
      url: targetUrl,
      openPage: targetPage
    }
  };
  event.waitUntil(
    self.registration.showNotification(data.title, options)
  );
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const targetUrl = event.notification && event.notification.data && event.notification.data.url
    ? event.notification.data.url
    : './index.html';
  event.waitUntil((async () => {
    const allClients = await clients.matchAll({ type: 'window', includeUncontrolled: true });
    if (allClients && allClients.length) {
      const client = allClients[0];
      if ('focus' in client) await client.focus();
      if ('navigate' in client) await client.navigate(targetUrl);
      return;
    }
    await clients.openWindow(targetUrl);
  })());
});

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)));
    await self.clients.claim();
  })());
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  const requestUrl = new URL(event.request.url);
  const isSameOrigin = requestUrl.origin === self.location.origin;
  if (isHtmlRequest(event.request)) {
    event.respondWith(networkFirst(event.request));
    return;
  }

  if (isSameOrigin) {
    event.respondWith(staleWhileRevalidate(event.request));
  }
});
