const CACHE = 'inventario-v2';
const APP_ASSETS = [
  './', 
  './index.html', 
  './style.css', 
  './app.js', 
  './manifest.json', 
  './icon.svg', 
  './lib/xlsx.full.min.js',
  './lib/zxing-library.min.js',
  './lib/zxing-browser.min.js'
  // 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(APP_ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request).then(response => {
      const copy = response.clone();
      if (response.ok || response.type === 'opaque') caches.open(CACHE).then(cache => cache.put(event.request, copy));
      return response;
    }).catch(() => caches.match('./index.html')))
  );
});
