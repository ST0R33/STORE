// IMPORTANTE: bumpar esta versão A CADA deploy que mude HTML/config.js.
// Sem o bump, o `activate` não limpa o cache antigo e o usuário fica preso na versão velha.
const CACHE = 'orahbuy-v7';
const SHELL = ['./index.html', './admin.html', './gestao_unificada.html', './config.js', './manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {})
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
  const req = e.request;
  const url = req.url;

  // GAS API e Cloudinary: sempre rede (dados frescos)
  if (url.includes('script.google.com') || url.includes('cloudinary.com')) {
    return;
  }

  // HTML + config.js: network-first. Antes era cache-first pra tudo, e isso travava
  // admin/gestao/index numa versão antiga mesmo depois do deploy (nem F5 resolvia — o SW
  // interceptava e devolvia o cache). config.js também nunca pode ficar velho: ele carrega
  // GAS_URL, e uma URL antiga em cache aponta pro deployment errado do Apps Script.
  const isHTML = req.mode === 'navigate' || (req.headers.get('accept') || '').includes('text/html');
  if (isHTML || url.includes('config.js')) {
    e.respondWith(
      fetch(req).then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
        return res;
      }).catch(() => caches.match(req).then(c => c || caches.match('./index.html')))
    );
    return;
  }

  // Demais assets (imagens, ícones): cache-first, fallback rede
  e.respondWith(
    caches.match(req).then(cached => cached || fetch(req))
  );
});
