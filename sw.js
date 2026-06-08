/**
 * Service Worker fuer BS pic upload
 * Version 5.0 - Cache-Busting fuer PDA-Kompatibilitaet
 */

const CACHE_NAME = 'bs-pic-upload-v5.0';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './app.js?v=4',
  './manifest.json'
];

// Installation: Cache statische Ressourcen + skipWaiting
self.addEventListener('install', event => {
  console.log('[SW] Installing v5.0...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[SW] Caching app shell');
        return cache.addAll(ASSETS_TO_CACHE);
      })
      .then(() => {
        console.log('[SW] Install complete, skipping waiting');
        return self.skipWaiting();
      })
      .then(() => {
        // Alle Clients ueber neue Version informieren
        return self.clients.matchAll({ includeUncontrolled: true }).then(clients => {
          clients.forEach(client => {
            client.postMessage({ type: 'SW_UPDATED', version: CACHE_NAME });
          });
        });
      })
  );
});

// Aktivierung: Alle alten Caches loeschen + clients.claim()
self.addEventListener('activate', event => {
  console.log('[SW] Activating v5.0...');
  event.waitUntil(
    caches.keys()
      .then(cacheNames => {
        return Promise.all(
          cacheNames
            .filter(name => name !== CACHE_NAME)
            .map(name => {
              console.log('[SW] Deleting old cache:', name);
              return caches.delete(name);
            })
        );
      })
      .then(() => {
        console.log('[SW] Claiming clients');
        return self.clients.claim();
      })
  );
});

// Fetch: Stale-While-Revalidate fuer statische Assets
// Fuer API-Requests: Network-first
self.addEventListener('fetch', event => {
  const { request } = event;
  const url = new URL(request.url);

  // Skip nicht-GET Requests
  if (request.method !== 'GET') {
    return;
  }

  // Microsoft Graph API und Login: Network-first, kein Caching
  if (url.hostname === 'graph.microsoft.com' || url.hostname === 'login.microsoftonline.com') {
    event.respondWith(
      fetch(request)
        .catch(() => {
          return caches.match(request).then(response => {
            return response || new Response('Offline - API nicht verfuegbar', {
              status: 503,
              statusText: 'Service Unavailable'
            });
          });
        })
    );
    return;
  }

  // Statische Assets: Stale-While-Revalidate
  // Gibt gecachte Version sofort zurueck UND aktualisiert Cache im Hintergrund
  event.respondWith(
    caches.open(CACHE_NAME).then(cache => {
      return cache.match(request).then(cachedResponse => {
        // Netzwerk-Request im Hintergrund starten
        const networkFetch = fetch(request).then(networkResponse => {
          if (networkResponse && networkResponse.status === 200) {
            cache.put(request, networkResponse.clone());
          }
          return networkResponse;
        }).catch(() => null);

        // Sofort cached Version zurueckgeben (falls vorhanden), sonst auf Netzwerk warten
        if (cachedResponse) {
          return cachedResponse;
        }
        return networkFetch.then(response => {
          return response || new Response('Offline - Ressource nicht verfuegbar', {
            status: 503,
            statusText: 'Service Unavailable'
          });
        });
      });
    })
  );
});
