// Service worker del Portal de Tecnicos. Cachea la app para que abra
// sin conexion despues de la primera visita. No depende de ningun
// recurso externo (a diferencia de Supervisor, este portal no usa CDN).

const CACHE_VERSION = 'portal-tecnicos-v3';
const NUCLEO = ['./index.html', './manifest.json', './logo-cobra.png', './icon-192.png', './icon-512.png', './fondo-tecnicos.jpg', './logo-tigo.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(NUCLEO)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((nombres) =>
      Promise.all(nombres.filter((n) => n !== CACHE_VERSION).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  event.respondWith(
    fetch(req)
      .then((res) => {
        const copia = res.clone();
        caches.open(CACHE_VERSION).then((cache) => cache.put(req, copia));
        return res;
      })
      .catch(() => caches.match(req))
  );
});
