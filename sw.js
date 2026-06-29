// IBI Settlement Tracker — service worker
// HTML/app-shell = NETWORK-FIRST (a normal refresh always gets the latest version;
// falls back to cache only when offline). Static assets = cache-first.
const CACHE = 'ibi-settlement-v5_4';
const ASSETS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS).catch(() => {})).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;

  const isHTML = req.mode === 'navigate' ||
                 (req.headers.get('accept') || '').includes('text/html');

  if (isHTML) {
    // NETWORK-FIRST: fetch fresh index.html, cache a copy, fall back to cache offline.
    e.respondWith(
      fetch(req)
        .then(resp => {
          try { const cp = resp.clone(); caches.open(CACHE).then(c => c.put('./index.html', cp)); } catch (_) {}
          return resp;
        })
        .catch(() => caches.match('./index.html').then(r => r || caches.match('./')))
    );
    return;
  }

  // Static assets: cache-first with runtime caching.
  e.respondWith(
    caches.match(req).then(cached => {
      if (cached) return cached;
      return fetch(req).then(resp => {
        try { const cp = resp.clone(); caches.open(CACHE).then(c => c.put(req, cp)); } catch (_) {}
        return resp;
      });
    })
  );
});
