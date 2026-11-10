// Guarda o app no aparelho para funcionar sem internet.
// Com internet, sempre busca a versão mais nova primeiro.
const CACHE = 'notas-v9';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL.map(u => new Request(u, { cache: 'reload' })))));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET' || new URL(r.url).origin !== location.origin) return;
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    try {
      const ctl = new AbortController();
      const t = setTimeout(() => ctl.abort(), 4000);
      const res = await fetch(r, { cache: 'no-cache', signal: ctl.signal });
      clearTimeout(t);
      if (res.ok) c.put(r, res.clone());
      return res;
    } catch (_) {
      return (await c.match(r, { ignoreSearch: true })) || (await c.match('index.html')) || Response.error();
    }
  })());
});
