// IMARAT Finance Dashboard — Service Worker v18
// Cache-busting: does NOT cache finance.html to ensure updates are always live

var CACHE_NAME = 'imarat-finance-v18';

// Install: skip waiting immediately
self.addEventListener('install', function(e) {
  self.skipWaiting();
});

// Activate: clear ALL old caches
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.map(function(k) { return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

// Fetch: ALWAYS go to network for app files and dynamic endpoints
// Only cache static assets like the Firebase SDK
self.addEventListener('fetch', function(e) {
  var url = e.request.url;

  // Never cache — always fetch fresh from network
  if (url.indexOf('finance.html') >= 0 ||
      url.indexOf('sw.js') >= 0 ||
      url.indexOf('manifest.json') >= 0 ||
      url.indexOf('script.google.com') >= 0 ||
      url.indexOf('firebaseapp.com') >= 0 ||
      url.indexOf('googleapis.com') >= 0) {
    return;
  }

  // Everything else (Firebase SDK CDN files etc) — serve from cache if available
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request);
    })
  );
});
