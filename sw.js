// IMARAT Finance Dashboard — Service Worker
// Handles offline caching so the chairman can open the app without internet

var CACHE_NAME = 'imarat-finance-v5';
var SHELL_URLS = [
  '/imarat-sap/finance.html',
  '/imarat-sap/manifest.json'
];

// Install: cache the app shell
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(SHELL_URLS);
    })
  );
  self.skipWaiting();
});

// Activate: remove old caches
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_NAME; })
            .map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Fetch: cache-first for app shell, network-first for data API calls
self.addEventListener('fetch', function(e) {
  var url = e.request.url;

  // Never intercept Apps Script / Firebase API calls — always go network
  if (url.indexOf('script.google.com') >= 0 ||
      url.indexOf('firebaseapp.com') >= 0 ||
      url.indexOf('googleapis.com') >= 0 ||
      url.indexOf('identitytoolkit') >= 0) {
    return; // Fall through to browser default (network)
  }

  // App shell: cache-first
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      return cached || fetch(e.request).then(function(response) {
        // Cache valid responses
        if (response && response.status === 200 && response.type === 'basic') {
          var clone = response.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(e.request, clone);
          });
        }
        return response;
      });
    })
  );
});
