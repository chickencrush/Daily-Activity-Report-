const CACHE_NAME = 'daily-report-cache-v2';
const ASSETS = ['./', './index.html', './manifest.json', './chicken_mascot.png'];

self.addEventListener('install', (event) => {
  // Aset yang belum ada tidak boleh menggagalkan instalasi
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => Promise.allSettled(ASSETS.map((a) => cache.add(a)))));
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return; // API Apps Script tidak di-cache
  // Jaringan dulu (agar versi baru langsung terpakai), cache sebagai cadangan offline
  event.respondWith(
    fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE_NAME).then((c) => c.put(req, copy));
      return res;
    }).catch(() => caches.match(req))
  );
});
