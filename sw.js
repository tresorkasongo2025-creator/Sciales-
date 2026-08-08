// Service Worker E-SCIALES UNILU+
const CACHE = 'esciales-v1';
const PRECACHE = [
  '/',
  '/static/css/style.css',
  '/static/logo-sciences-sociales.png',
  '/static/croix-rouge.png'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(PRECACHE)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Network-first pour les requêtes POST et les pages dynamiques
  if (e.request.method !== 'GET') return;
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
