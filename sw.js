// IMARAT Finance Dashboard — Service Worker v10
// Cache-busting version: does NOT cache finance.html to ensure updates are always live

var CACHE_NAME = 'imarat-finance-v15';

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

// Fetch: ALWAYS go to network for finance.html and sw.js
// Only cache static assets like Firebase SDK
self.addEventListener('fetch', function(e) {
  var url = e.request.url;
  
  // Never cache the main app files — always fetch fresh
  if (url.indexOf('finance.html') >= 0 ||
      url.indexOf('sw.js') >= 0 ||
      url.indexOf('manifest.json') >= 0 ||
      url.indexOf('script.google.com') >= 0 ||
      url.indexOf('firebaseapp.com') >= 0 ||
      url.indexOf('googleapis.com') >= 0) {
    return; // Always fetch from network
  }
  
  // For everything else (Firebase SDK etc), use cache
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request);
    })
  );
});
