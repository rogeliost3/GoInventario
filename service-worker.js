const CACHE = 'inventario-v3';
const FILES = [
  './', 
  './index.html', 
  './style.css', 
  './app.js', 
  './manifest.json', 
  './icon.svg', 
  './lib/xlsx.full.min.js',
  './lib/barcode-detector.iife.js',
  './lib/zxing_reader.wasm'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE)
      // 'reload' evita que se guarde una copia vieja de la caché HTTP del navegador
      .then(cache => cache.addAll(FILES.map(url => new Request(url, { cache: 'reload' }))))
      .then(() => self.skipWaiting())          // no esperar al cierre de la app
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())        // tomar el control de la app ya abierta
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(caches.match(event.request).then(r => r || fetch(event.request)));
});

// self.addEventListener('install', event => {
//   event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(APP_ASSETS)).then(() => self.skipWaiting()));
// });

// self.addEventListener('activate', event => {
//   event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
// });

// self.addEventListener('fetch', event => {
//   if (event.request.method !== 'GET') return;
//   event.respondWith(
//     caches.match(event.request).then(cached => cached || fetch(event.request).then(response => {
//       const copy = response.clone();
//       if (response.ok || response.type === 'opaque') caches.open(CACHE).then(cache => cache.put(event.request, copy));
//       return response;
//     }).catch(() => caches.match('./index.html')))
//   );
// });
