// ============================================================
//   SERVICE WORKER — AbsenSiswa PWA
// ============================================================

var CACHE_NAME = 'absensiswa-v1';
var ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

// Install: cache semua aset
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

// Activate: hapus cache lama
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_NAME; })
            .map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Fetch: cache-first untuk aset, network-first untuk API
self.addEventListener('fetch', function(e) {
  var url = e.request.url;

  // Selalu ke network untuk Apps Script API
  if (url.includes('script.google.com')) {
    e.respondWith(
      fetch(e.request).catch(function() {
        return new Response(
          JSON.stringify({status:'error', message:'Tidak ada koneksi internet'}),
          { headers: { 'Content-Type': 'application/json' } }
        );
      })
    );
    return;
  }

  // Cache-first untuk aset lokal
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request).then(function(response) {
        // Simpan ke cache jika berhasil
        if (response && response.status === 200 && response.type === 'basic') {
          var clone = response.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(e.request, clone);
          });
        }
        return response;
      });
    }).catch(function() {
      // Offline fallback
      if (e.request.destination === 'document') {
        return caches.match('./index.html');
      }
    })
  );
});
