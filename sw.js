// Zwischenspeicher, damit KISY auch ohne Netz startet (z. B. auf dem Golfplatz).
// Programmdateien: zuerst aus dem Netz (immer aktuell), bei fehlendem Netz aus dem Zwischenspeicher.
// Die Ordner-Zugriffe am PC (/lokal/) und Anfragen an Google werden nie zwischengespeichert.
const CACHE = "kisy-v3";
const DATEIEN = ["./", "index.html", "style.css", "manifest.webmanifest", "icon-180.png", "icon-192.png", "vendor/xlsx.full.min.js",
  "js/app.js", "js/grundlagen.js", "js/daten.js", "js/excel.js", "js/speicher.js", "js/grafik.js",
  "js/abgleich.js", "js/google.js", "js/konfig.js"];

self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(DATEIEN)).then(() => self.skipWaiting())); });
// Nur eigene alte Zwischenspeicher löschen: andere Apps unter derselben Adresse (github.io) haben ihre eigenen.
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(n => n.startsWith("kisy-") && n !== CACHE).map(n => caches.delete(n)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== self.location.origin || url.pathname.includes("/lokal/")) return;
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    try {
      const r = await fetch(e.request);
      if (r.ok) c.put(e.request, r.clone());
      return r;
    } catch (err) {
      const alt = await c.match(e.request, { ignoreSearch: true });
      if (alt) return alt;
      throw err;
    }
  })());
});
