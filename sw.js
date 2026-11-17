// Cache simples da interface para abrir o app sem rede (os dados nunca sao cacheados).
const CACHE = 'ptv-v11';
const ARQS = ['./', 'index.html', 'manifest.webmanifest', 'logo-mark.png', 'logo-word.png', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png', 'favicon.png', 'favicon-32.png'];
self.addEventListener('install', (e) => { e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ARQS))); self.skipWaiting(); });
self.addEventListener('activate', (e) => { e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))); self.clients.claim(); });
self.addEventListener('fetch', (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then((r) => { const cp = r.clone(); caches.open(CACHE).then((c) => c.put(e.request, cp)); return r; }).catch(() => caches.match(e.request)));
});
