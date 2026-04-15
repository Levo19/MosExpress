const CACHE_NAME = 'mosexpress-cache-v44';
const urlsToCache = [
  './',
  './index.html',
  './manifest.json',
  'https://unpkg.com/vue@3.4.21/dist/vue.global.prod.js',
  'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(async cache => {
        console.log('Opened cache');
        try { await cache.addAll(urlsToCache); } catch(e) { console.warn('SW cache install parcial:', e); }
      })
  );
  self.skipWaiting();
});

self.addEventListener('fetch', event => {
  // Solo interceptamos peticiones GET (no interceptamos POST al GAS)
  if (event.request.method !== 'GET') return;
  
  // No interceptar peticiones a la API del backend
  if (event.request.url.includes('script.google.com') || event.request.url.includes('api.printnode.com')) {
      return;
  }

  event.respondWith(
    caches.match(event.request)
      .then(response => {
        // Cache hit - return response
        if (response) {
          return response;
        }

        // Clone the request because it's a stream and can only be consumed once
        var fetchRequest = event.request.clone();

        return fetch(fetchRequest).then(
          function(response) {
            if(!response || response.status !== 200) {
              return response;
            }
            // Cachear respuestas básicas y CORS válidas
            if(response.type !== 'basic' && response.type !== 'cors') {
              return response;
            }

            var responseToCache = response.clone();

            // Usar waitUntil para garantizar que el cache se guarda antes de que el SW se duerma
            event.waitUntil(
              caches.open(CACHE_NAME).then(function(cache) {
                if (event.request.url.startsWith('http')) {
                    return cache.put(event.request, responseToCache);
                }
              })
            );

            return response;
          }
        ).catch(function(error) {
            console.error('Fetch event failed:', error);
            return Response.error();
        });
      })
  );
});

self.addEventListener('activate', event => {
  const cacheWhitelist = [CACHE_NAME];
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheWhitelist.indexOf(cacheName) === -1) {
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  self.clients.claim();
});
