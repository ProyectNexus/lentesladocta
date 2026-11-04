/* Lentes La Docta · service worker
   - Páginas: primero internet (siempre la última versión); sin conexión, la guardada.
   - Fotos, íconos y estilos: se guardan para que abra rápido.
   - Los datos (Google Apps Script) nunca se guardan acá: siempre van en vivo. */
const VERSION = '2026-11-03-1';
const CACHE = 'lld-' + VERSION;
const BASE = [
  './', './index.html', './config.js', './manifest.webmanifest',
  './panel/', './panel/index.html', './panel/manifest.webmanifest',
  './icons/icon-192.png', './icons/favicon.svg', './panel/icons/icon-192.png',
  './fotos/aviador_marron.webp',
  './fotos/clubmaster.webp',
  './fotos/redondo_carey.webp',
  './fotos/aviador_azul.webp',
  './fotos/carey_oro.webp',
  './fotos/clubmaster_receta.webp',
  './fotos/corazon.webp',
  './fotos/negros.webp',
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(BASE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k.startsWith('lld-') && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;               // Apps Script, Google Fonts, Drive: directo
  const esPagina = req.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname.endsWith('config.js');
  if (esPagina) {
    e.respondWith(fetch(req).then(r => {
      const copia = r.clone(); caches.open(CACHE).then(c => c.put(req, copia)); return r;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match('./index.html'))));
    return;
  }
  e.respondWith(caches.match(req).then(r => r || fetch(req).then(res => {
    const copia = res.clone(); caches.open(CACHE).then(c => c.put(req, copia)); return res;
  })));
});
