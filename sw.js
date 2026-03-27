// PezzaliApp — Cormach Dashboard Service Worker v1774620808
// Cache disabilitata — passa-through diretto
self.addEventListener('install', e => { self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.map(k => caches.delete(k))))
  );
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  // Nessuna cache — vai sempre alla rete
  e.respondWith(fetch(e.request).catch(() => new Response('Offline', {status: 503})));
});
