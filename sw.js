// Angkor Café POS — offline support
// Page: network-first (always get the newest version when online, fall back to cache offline)
// CDN assets (Tailwind, Font Awesome, Chart.js, Google Fonts): cache-first
const CACHE = 'cafe-pos-v1';
const CORE = ['./', './index.html', './manifest.json', './icon.svg'];
const CDN_HOSTS = ['cdn.tailwindcss.com', 'cdnjs.cloudflare.com', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', e => {
    e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
    e.waitUntil(caches.keys()
        .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
        .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
    const req = e.request;
    if (req.method !== 'GET') return; // never touch Google Sheet sync (POST)
    const url = new URL(req.url);

    if (url.origin === location.origin) {
        e.respondWith(
            fetch(req).then(res => {
                if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
                return res;
            }).catch(() => caches.match(req).then(r => r || caches.match('./index.html')))
        );
        return;
    }

    if (CDN_HOSTS.includes(url.hostname)) {
        e.respondWith(
            caches.match(req).then(hit => hit || fetch(req).then(res => {
                if (res.ok || res.type === 'opaque') { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
                return res;
            }))
        );
    }
});
