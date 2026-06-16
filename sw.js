var CACHE = 'cts-v12';
var ASSETS = ['./', './index.html', './manifest.json', './icons/icon.svg'];

self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(c) { return c.addAll(ASSETS); })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.filter(function(k) { return k !== CACHE; }).map(function(k) { return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e) {
  // ไม่ cache GAS API calls (รวม redirect target ที่ googleusercontent.com)
  if (e.request.url.includes('script.google.com') || e.request.url.includes('googleapis.com') || e.request.url.includes('googleusercontent.com')) return;
  // ไม่ cache Google Fonts (ถ้า offline จะใช้ fallback font)
  if (e.request.url.includes('fonts.googleapis') || e.request.url.includes('fonts.gstatic')) return;

  e.respondWith(
    caches.match(e.request).then(function(cached) {
      var network = fetch(e.request).then(function(res) {
        if (res.ok && e.request.method === 'GET') {
          var clone = res.clone();
          caches.open(CACHE).then(function(c) { c.put(e.request, clone); });
        }
        return res;
      }).catch(function() { return cached || new Response('Offline', {status: 503}); });
      return cached || network;
    })
  );
});
