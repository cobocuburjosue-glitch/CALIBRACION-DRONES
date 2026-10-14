/* Service worker: deja la app guardada en el teléfono para abrirla sin internet.
   Los envíos a Google (Apps Script) NO pasan por aquí: esos los maneja la cola de la app. */
const CACHE = 'calibracion-v1';
const PRECACHE = [
  './', 'index.html', 'manifest.json', 'logo.png', 'icon-192.png', 'icon-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
  'https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(PRECACHE.map(u =>
    c.add(new Request(u, u.startsWith('http') ? { mode: 'no-cors' } : {})).catch(() => {})
  ))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

function conTiempo(p, ms) {
  return new Promise((ok, ko) => { const t = setTimeout(() => ko(new Error('timeout')), ms); p.then(r => { clearTimeout(t); ok(r); }, e => { clearTimeout(t); ko(e); }); });
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (/script\.google(usercontent)?\.com$/.test(url.hostname) || /googleusercontent\.com$/.test(url.hostname)) return; // API: siempre red
  if (/\.apk$|version\.json$/.test(url.pathname)) return;

  // Página: primero internet (para recibir mejoras); si no hay o tarda, la guardada.
  if (req.mode === 'navigate') {
    e.respondWith((async () => {
      const c = await caches.open(CACHE);
      try {
        const r = await conTiempo(fetch(req), 5000);
        if (r && r.ok) c.put('index.html', r.clone());
        return r;
      } catch (_) {
        return (await c.match('index.html')) || (await c.match('./')) || (await caches.match(req)) || Response.error();
      }
    })());
    return;
  }

  // Archivos (logo, fuentes, gráficas): la copia guardada al instante y se actualiza por detrás.
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    const guardada = await c.match(req);
    const red = fetch(req).then(r => { if (r && (r.ok || r.type === 'opaque')) c.put(req, r.clone()); return r; }).catch(() => guardada);
    return guardada || red;
  })());
});
