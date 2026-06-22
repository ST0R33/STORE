const CACHE = 'orahbuy-v2';
const SHELL = ['./index.html', './config.js'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(SHELL))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = e.request.url;

  // GAS API e Cloudinary: sempre rede (dados frescos)
  if (url.includes('script.google.com') || url.includes('cloudinary.com')) {
    return;
  }

  // Shell HTML/JS: cache primeiro, fallback rede
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});
