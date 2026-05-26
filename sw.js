const CACHE = 'coc-v2.4';
const STATIC = ['/', '/index.html', '/manifest.json', '/icon.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATIC)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = e.request.url;
  // Skip non-http requests (chrome-extension etc) and POST/non-GET requests
  if (!url.startsWith('http') || e.request.method !== 'GET') return;
  // Network only for API calls
  if (url.includes('supabase.co') || url.includes('fonts.googleapis') || url.includes('fonts.gstatic')) {
    e.respondWith(fetch(e.request).catch(() => new Response('', {status: 503})));
    return;
  }
  // Cache first for static assets
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (!res || res.status !== 200 || res.type === 'opaque') return res;
        const clone = res.clone();
        caches.open(CACHE).then(c => { try { c.put(e.request, clone); } catch(err) {} });
        return res;
      }).catch(() => caches.match('/index.html'));
    })
  );
});
