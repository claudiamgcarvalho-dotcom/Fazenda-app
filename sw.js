const CACHE_NAME = 'boivirtual-v2';
const SHELL = [
  '/Fazenda-app/',
  '/Fazenda-app/index.html',
  '/Fazenda-app/FuturaStd-Light.otf',
  '/Fazenda-app/FuturaStd-Book.otf',
  '/Fazenda-app/FuturaStd-Medium.otf',
  '/Fazenda-app/FuturaStd-Bold.otf',
  '/Fazenda-app/icon-192.png',
  '/Fazenda-app/icon-512.png',
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  if (e.request.url.includes('script.google.com')) return;
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});
