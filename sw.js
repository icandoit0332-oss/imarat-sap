// IMARAT — shared Service Worker (serves index.html, team.html, finance.html)
// Strategy: network-first for pages (so new deploys show immediately, cache for offline),
// cache-first for other same-origin assets, and NEVER touch cross-origin calls
// (Firebase, Apps Script, Google Fonts) so live data always hits the network.
// Bump CACHE_VERSION to force every client to drop old caches.

var CACHE_VERSION = 'imarat-pwa-v1';

self.addEventListener('install', function (e) {
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys()
      .then(function (keys) {
        return Promise.all(keys.map(function (k) {
          if (k !== CACHE_VERSION) return caches.delete(k);
        }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;                 // never cache writes

  var url;
  try { url = new URL(req.url); } catch (err) { return; }

  // Leave all cross-origin requests alone (Firebase / Apps Script / fonts / CDNs)
  if (url.origin !== self.location.origin) return;

  var isPage = req.mode === 'navigate' ||
    (req.headers.get('accept') || '').indexOf('text/html') !== -1;

  if (isPage) {
    // Network-first: fresh when online, fall back to the cached copy offline.
    e.respondWith(
      fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(CACHE_VERSION).then(function (c) { c.put(req, copy); });
        return res;
      }).catch(function () {
        return caches.match(req).then(function (r) { return r || caches.match(req.url); });
      })
    );
    return;
  }

  // Other same-origin assets: cache-first, then network (and cache the result).
  e.respondWith(
    caches.match(req).then(function (cached) {
      return cached || fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(CACHE_VERSION).then(function (c) { c.put(req, copy); });
        return res;
      }).catch(function () { return cached; });
    })
  );
});
