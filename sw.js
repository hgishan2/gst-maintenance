// Service worker: keeps the technician app working with no signal.
// Bump CACHE_VERSION whenever you change index.html / config.js so phones pick up the new files.
const CACHE_VERSION = 'gst-ops-v13';

const CORE = [
  './',
  './index.html',
  './config.js',
  './manifest.json',
  'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css',
  'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js',
  'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css',
  'https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap',
  'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/fonts/bootstrap-icons.woff2',
  'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/fonts/bootstrap-icons.woff'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then(cache =>
      Promise.all(CORE.map(url =>
        fetch(new Request(url, { cache: 'reload' }))
          .then(res => { if (res && (res.ok || res.type === 'opaque')) return cache.put(url, res); })
          .catch(() => {})                     // one missing file must not block the install
      ))
    ).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Own files: try the network first (so updates arrive), but fall back to the cache after 4 s or when offline.
function networkFirst(req) {
  return new Promise(resolve => {
    let settled = false;
    const done = r => { if (!settled) { settled = true; resolve(r); } };
    const timer = setTimeout(() => {
      caches.match(req, { ignoreSearch: true }).then(c => { if (c) done(c); });
    }, 4000);
    fetch(req).then(res => {
      clearTimeout(timer);
      if (res && res.ok) { const copy = res.clone(); caches.open(CACHE_VERSION).then(c => c.put(req, copy)); }
      done(res);
    }).catch(() => {
      clearTimeout(timer);
      caches.match(req, { ignoreSearch: true }).then(c => {
        if (c) return done(c);
        if (req.mode === 'navigate') return caches.match('./index.html').then(i => done(i || Response.error()));
        done(Response.error());
      });
    });
  });
}

// Libraries / fonts from the CDN: cache first (they never change for a pinned version).
function cacheFirst(req) {
  return caches.match(req, { ignoreSearch: true }).then(hit => {
    if (hit) return hit;
    return fetch(req).then(res => {
      if (res && (res.ok || res.type === 'opaque')) { const copy = res.clone(); caches.open(CACHE_VERSION).then(c => c.put(req, copy)); }
      return res;
    });
  });
}

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;                              // POSTs (saving data) always go straight to the network
  const url = new URL(req.url);
  if (url.hostname.endsWith('script.google.com') || url.hostname.endsWith('googleusercontent.com')) return;   // data calls: app handles offline itself
  if (url.origin === self.location.origin) event.respondWith(networkFirst(req));
  else event.respondWith(cacheFirst(req));
});
