/* Offline support: app shell is cached at install; libraries are cached on first use and warmed in the background. */
const VERSION = 'my-office-v0.1.1';
const SHELL = ['./', 'index.html', 'app.css', 'manifest.webmanifest', 'js/store.js', 'js/sync.js', 'js/core.js', 'js/docs.js', 'js/sheets.js', 'js/slides.js', 'js/pdfview.js', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/maskable-512.png'];
const VENDOR = ['vendor/mammoth.min.js', 'vendor/docx-kit.js', 'vendor/formula-engine.js', 'vendor/exceljs.min.js', 'vendor/pptxgen.bundle.js', 'vendor/jszip.min.js', 'vendor/pdf.js', 'vendor/pdf.worker.min.mjs'];
self.addEventListener('install', e => { e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('message', e => {
  if (e.data && e.data.type === 'warm') e.waitUntil(caches.open(VERSION).then(async c => { for (const u of VENDOR) { if (!(await c.match(u))) { try { await c.add(u); } catch (err) {} } } }));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  const isNav = e.request.mode === 'navigate';
  e.respondWith((async () => {
    const c = await caches.open(VERSION);
    if (isNav) { // network first for the page so updates arrive; cached copy when offline
      try { const r = await fetch(e.request); if (r.ok) c.put('index.html', r.clone()); return r; } catch (err) { return (await c.match('index.html')) || Response.error(); }
    }
    const hit = await c.match(e.request, {ignoreSearch: true});
    if (hit) return hit;
    const r = await fetch(e.request);
    if (r.ok && (u.pathname.includes('/vendor/') || u.pathname.includes('/js/') || u.pathname.includes('/icons/'))) c.put(e.request, r.clone());
    return r;
  })());
});
