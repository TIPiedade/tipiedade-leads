/* Ti'Piedade · Operações — service worker
   Guarda a aplicação para funcionar sem rede.
   Os dados (historico.json) são tratados pela própria app em localStorage. */
const V = 'tp-ops-v1';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png', './catalogo.html'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(SHELL.map(u => new Request(u, {cache:'reload'})))).then(() => self.skipWaiting()).catch(()=>self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;                       // escritas nunca são servidas da cache
  const u = new URL(r.url);
  if (u.pathname.endsWith('historico.json') || u.hostname.includes('workers.dev')) return;  // dados sempre frescos

  if (r.mode === 'navigate') {                          // páginas: rede primeiro, cache como rede de segurança
    e.respondWith(fetch(r).then(res => {
      const c = res.clone(); caches.open(V).then(x => x.put(r, c)); return res;
    }).catch(() => caches.match(r).then(m => m || caches.match('./index.html'))));
    return;
  }
  e.respondWith(caches.match(r).then(m => m || fetch(r).then(res => {
    if (res.ok && (u.origin === location.origin || u.hostname.includes('fonts.'))) {
      const c = res.clone(); caches.open(V).then(x => x.put(r, c));
    }
    return res;
  }).catch(() => m)));
});
