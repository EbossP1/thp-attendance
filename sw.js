// THP-Ghana Attendance — Service Worker
// ⚠ DEPLOY RULE: bump the version number below on EVERY deploy
// (v2 → v3 → v4 …). That one change makes all installed apps
// fetch fresh files and reload themselves automatically.
const CACHE_NAME = 'thp-attendance-v3.6';
const ASSETS = ['/', '/index.html', '/app.js', '/styles.css'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(ASSETS)).catch(()=>{}));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Network-first with revalidation (bypasses stale HTTP cache), cache fallback for offline
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request, { cache: 'no-cache' })
      .then(response => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then(c => c.put(event.request, clone)).catch(()=>{});
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
