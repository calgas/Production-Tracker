// Service worker — app-shell caching, with a hard guarantee that stock
// data itself is never served stale.
//
// Strategy:
//  - Anything hitting the Apps Script backend (script.google.com /
//    script.googleusercontent.com), and any non-GET request, is ALWAYS
//    fetched from the network and never cached — stock numbers must never
//    come from a cache.
//  - Everything else — the app shell (index.html, manifest.json) and the
//    static libraries it loads — is cached so a repeat visit is instant,
//    and an install with no signal shows the last-seen shell instead of a
//    blank white screen. Cache-first, with a background refetch to keep
//    the cache warm for next time.
//
// The sibling apps (ELE Tracker and the rest) are served from the same
// origin, calgas.github.io, and Cache Storage is shared across an origin,
// not per app. Every cache this worker owns carries CACHE_PREFIX, and
// cleanup only touches caches with that prefix — deleting "everything that
// isn't mine" would wipe each sibling's offline copy whenever this worker
// updated. A new sibling app must use a prefix of its own.
//
// Bump CACHE_VERSION whenever index.html / manifest.json / this file
// changes, so an already-installed client picks up the new shell instead
// of being stuck on an old cached one. Keep this in step with APP_VERSION
// in index.html.
const CACHE_PREFIX = 'production-tracker-shell-';
const CACHE_VERSION = CACHE_PREFIX + 'v1';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  'https://calgas.github.io/Stock-Management/logo/icon-192.png',
  'https://calgas.github.io/Stock-Management/logo/icon-512.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_VERSION)
      .then((cache) => cache.addAll(APP_SHELL))
      .catch(() => {}) // don't block install if a shell asset is briefly unreachable
  );
  // NOT skipWaiting() here. Taking over immediately would swap the shell
  // under someone mid-entry, and on a shop floor that means a half-typed
  // transaction disappearing. The new worker waits; the page notices it and
  // offers a Reload, and only then does the message below let it through.
});

self.addEventListener('message', (e) => {
  if (e.data && e.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((names) => Promise.all(
        names
          .filter((n) => n.startsWith(CACHE_PREFIX) && n !== CACHE_VERSION)
          .map((n) => caches.delete(n))
      ))
      .then(() => self.clients.claim()) // take control of already-open tabs
  );
});

function isBackendRequest(url) {
  return url.hostname === 'script.google.com' || url.hostname === 'script.googleusercontent.com';
}

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);

  // Live stock data — always network, never cached or served from cache.
  if (isBackendRequest(url) || e.request.method !== 'GET') {
    e.respondWith(fetch(e.request));
    return;
  }

  // App shell & static assets — cache-first, refresh in the background so
  // the next visit has the latest version too.
  e.respondWith(
    caches.open(CACHE_VERSION).then((cache) =>
      cache.match(e.request).then((cached) => {
        const network = fetch(e.request)
          .then((resp) => {
            if (resp && resp.ok) cache.put(e.request, resp.clone());
            return resp;
          })
          .catch(() => cached); // offline with nothing fresh — fall back to cache if we have it
        return cached || network;
      })
    )
  );
});
