const CACHE_NAME = 'energy-monitor-v1';
const PRECACHE_ASSETS = [
  './',
  'index.html',
  'manifest.json',
  'icon.jpg'
];

// Install Event: pre-cache the shell files
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[Service Worker] Pre-caching offline shell');
      return cache.addAll(PRECACHE_ASSETS);
    })
  );
});

// Activate Event: clean up old caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('[Service Worker] Removing old cache', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch Event: handle requests offline
self.addEventListener('fetch', (event) => {
  // Only handle local GET requests
  if (event.request.method !== 'GET') return;

  const url = new URL(event.request.url);

  // If it's an external resource (not on our domain), let it pass normally,
  // but we can cache fonts or CDN stylesheets if necessary.
  if (url.origin !== self.location.origin) {
    // Optional caching for google fonts/CDNs
    if (url.hostname.includes('fonts.googleapis.com') || url.hostname.includes('fonts.gstatic.com')) {
      event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
          if (cachedResponse) return cachedResponse;
          return fetch(event.request).then((networkResponse) => {
            const responseToCache = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, responseToCache);
            });
            return networkResponse;
          });
        })
      );
    }
    return;
  }

  // Handle local navigation requests (e.g. typing a path, or clicking refresh)
  // Serve /index.html as the shell for SPA routing
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request).catch(() => {
        console.log('[Service Worker] Offline: serving index.html shell');
        return caches.match('index.html') || caches.match('./');
      })
    );
    return;
  }

  // For all other local assets, use a Cache-First strategy with Network Fallback
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        // Serve immediately from cache for offline responsiveness
        // Try updating cache in background (Stale-While-Revalidate)
        fetch(event.request).then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, networkResponse);
            });
          }
        }).catch(() => {
          // Fail silently in background if offline
        });
        return cachedResponse;
      }

      // If not cached, fetch from network
      return fetch(event.request).then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200) {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      }).catch((err) => {
        console.warn('[Service Worker] Fetch failed offline for:', event.request.url, err);
        // Avoid returning undefined to prevent browser app load crashes
        if (event.request.destination === 'image') {
          return caches.match('icon.jpg') || new Response('', { status: 404 });
        }
        return new Response('Автономный режим: ресурс недоступен', { 
          status: 503, 
          statusText: 'Service Unavailable' 
        });
      });
    })
  );
});
