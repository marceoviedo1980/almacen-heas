// Prefijo propio: al actualizar solo se borran las cachés de ESTA app.
// Otras PWA del mismo dominio (marceoviedo1980.github.io) no se tocan.
const PREFIX = 'almacen-heas-';
// Cambia este número cada vez que subas un almacen.json nuevo
const CACHE = PREFIX + 'v3';
const FILES = ['./index.html', './almacen.json', './manifest.webmanifest', './icon-192.png', './icon-512.png', './logo-heas.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k.startsWith(PREFIX) && k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

// Red primero (saldos frescos si hay internet); si no hay conexión, la copia guardada
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    fetch(e.request, { cache: 'no-store' })
      .then(res => {
        if (res && res.ok) { const clone = res.clone(); caches.open(CACHE).then(c => c.put(e.request, clone)); }
        return res;
      })
      .catch(() => caches.open(CACHE).then(c => c.match(e.request, { ignoreSearch: true })))
  );
});
