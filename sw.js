
// IMARAT — shared Service Worker (serves index.html, team.html, finance.html, malls.html)
// Strategy: network-first for pages (new deploys show immediately; cache only for true offline),
// cache-first for other same-origin assets, and NEVER touch cross-origin calls
// (Firebase, Apps Script, Google Fonts) so live data always hits the network.
// Bump CACHE_VERSION on EVERY deploy to force every client to drop old caches.
var CACHE_VERSION = 'imarat-pwa-v3';
 
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
 
// Network-first with a hard timeout: if the network is slow (not offline),
// we still wait briefly, then fall back to cache — but we never PREFER stale.
function fetchPageFresh(req) {
  return new Promise(function (resolve, reject) {
    var settled = false;
    var timer = setTimeout(function () {
      if (!settled) { settled = true; reject(new Error('timeout')); }
    }, 4000);
    fetch(req).then(function (res) {
      if (settled) return;
      settled = true; clearTimeout(timer); resolve(res);
    }).catch(function (err) {
      if (settled) return;
      settled = true; clearTimeout(timer); reject(err);
    });
  });
}
 
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
    // Network-first: fresh when online, cached copy only as offline fallback.
    e.respondWith(
      fetchPageFresh(req).then(function (res) {
        // Only cache good responses (don't cache 4xx/5xx/opaque errors)
        if (res && res.ok) {
          var copy = res.clone();
          caches.open(CACHE_VERSION).then(function (c) { c.put(req, copy); });
        }
        return res;
      }).catch(function () {
        return caches.match(req).then(function (r) { return r || caches.match(req.url); });
      })
    );
    return;
  }
 
  // Other same-origin assets: cache-first, then network (and cache good results).
  e.respondWith(
    caches.match(req).then(function (cached) {
      return cached || fetch(req).then(function (res) {
        if (res && res.ok) {
          var copy = res.clone();
          caches.open(CACHE_VERSION).then(function (c) { c.put(req, copy); });
        }
        return res;
      }).catch(function () { return cached; });
    })
  );
});
