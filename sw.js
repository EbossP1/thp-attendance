// THP-Ghana Attendance — Service Worker
// ⚠ DEPLOY RULE: bump the version number below on EVERY deploy
// (v2 → v3 → v4 …). That one change makes all installed apps
// fetch fresh files and reload themselves automatically.
const CACHE_NAME = 'thp-attendance-v5.0';
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

// Only the app's own files are handled here.
// Supabase, Apps Script, fonts and CDNs go straight to the network:
// caching them slowed every data load and stored staff data and
// session tokens on the device.
self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  if (new URL(req.url).origin !== self.location.origin) return;   // cross-origin → browser handles it
  if (req.headers.has('range')) return;                            // video streaming (farmers.mp4)

  // Network-first with revalidation, cache fallback for offline
  event.respondWith(
    fetch(req, { cache: 'no-cache' })
      .then(response => {
        if (response.ok && response.type === 'basic') {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(c => c.put(req, clone)).catch(()=>{});
        }
        return response;
      })
      .catch(() => caches.match(req))
  );
});
