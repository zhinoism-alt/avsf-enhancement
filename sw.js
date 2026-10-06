const CACHE = 'avsf-shell-v1';
const SHELL = ['./', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// Network first for the page itself (always get the latest dashboard), cache as offline fallback.
// Static CDN assets: cache first. Firebase / API calls are never intercepted.
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const staticCdn = /(^|\.)(cdnjs\.cloudflare\.com|gstatic\.com|googleapis\.com|jsdelivr\.net)$/.test(url.hostname)
    && !/firestore|identitytoolkit|securetoken/.test(url.hostname);
  if (url.origin === location.origin) {
    e.respondWith(fetch(req).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(req, copy));
      return res;
    }).catch(() => caches.match(req).then(r => r || caches.match('./'))));
  } else if (staticCdn && /\.(js|css|woff2?)$|fonts\.googleapis\.com/.test(url.href)) {
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(req, copy));
      return res;
    })));
  }
});
