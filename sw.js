// Orders & Dispatch moved to https://calgas.github.io/Orders-Dispatch/
// (27 Nov 2026). This replaces the old app's service worker at
// /Production-Tracker/: it removes the old app's cached copy, sends any
// open window to the new address, and then removes itself.
const NEW_URL = 'https://calgas.github.io/Orders-Dispatch/';
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter((n) => n.startsWith('production-tracker-shell-')).map((n) => caches.delete(n)));
    await self.clients.claim();
    const wins = await self.clients.matchAll({ type: 'window' });
    await Promise.all(wins.map((c) => { const u = new URL(c.url); return c.navigate(NEW_URL + u.search + u.hash).catch(() => {}); }));
    await self.registration.unregister();
  })());
});
self.addEventListener('fetch', () => {}); // nothing is served from here
