const CACHE_NAME = 'formsmobile-v51';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './css/styles.css',
  './js/app.js',
  './js/store.js',
  './manifest.json'
];

self.addEventListener('install', (event) => {
  self.skipWaiting(); // Force activate immediately
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim(); // Take control of all pages immediately
});

self.addEventListener('fetch', (event) => {
  // Network first, fallback to cache (strip query params for cache matching)
  event.respondWith(
    fetch(event.request).then((response) => {
      return response;
    }).catch(() => {
      // Strip query string for cache lookup
      const url = new URL(event.request.url);
      url.search = '';
      return caches.match(url.toString()) || caches.match(event.request);
    })
  );
});
